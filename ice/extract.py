"""Read an Ice Monitoring whiteboard photo with Claude and return draft rows.

The result is only a draft: the user always reviews and corrects it on the
review screen before anything is written to Tracker.xlsx.
"""
from __future__ import annotations

import base64
import difflib
import os
from dataclasses import dataclass, field
from io import BytesIO

from PIL import Image, ImageOps
from pydantic import BaseModel, Field

from .tracker import Client, match_client

MAX_SIDE = 2400  # px; large phone photos are downscaled before sending


class BoardEntry(BaseModel):
    rider: str = Field(description="Rider name at the top of the column this entry is under")
    client: str = Field(description="Client name exactly as written on the board")
    quantity: float = Field(description="Quantity written before the client name; 1/2 = 0.5")
    unclear: bool = Field(description="True if the handwriting for this entry was hard to read")
    note: str = Field(description="Short note on what was unclear, or empty string")


class BoardReading(BaseModel):
    entries: list[BoardEntry]
    riders_off: list[str] = Field(description="Riders marked OFF or with no entries")
    left_out: list[str] = Field(
        description="Marks that look like entries but don't fit the 'quantity client' pattern, quoted as written"
    )
    notes: list[str] = Field(description="Anything else the reviewer should check")


SYSTEM_PROMPT = """You read photos of the Ice Monitoring whiteboard used by an ice delivery business and transcribe the day's deliveries.

How the board is laid out:
- The board is arranged in columns. The name at the top of each column is the rider. Every entry below it belongs to that rider until the next rider's name.
- Each entry is a quantity followed by a client name, e.g. "2 Lisa" means quantity 2 for client Lisa. "1/2" or a stacked fraction means 0.5.
- "OFF" written under a rider means that rider had no sales.
- Ignore stray marks, dashes, tallies and anything that is not a rider or an entry.

Rules:
- Read the image itself carefully, column by column. Zoom in mentally on unclear parts before deciding.
- Never invent a quantity. If you cannot read a quantity, put the entry in left_out instead of entries.
- If handwriting is unclear, give your best reading, set unclear=true and say what was unclear in note.
- Write client names as they appear on the board; if one clearly matches a name in the known client list, use the list's spelling.
- Use the known rider names when a column header is an obvious misspelling of one (e.g. "Jomor" = Jomar, "Pick-up" = Pickup)."""


@dataclass
class DraftRow:
    rider: str
    client: str
    qty: float
    flags: list[str] = field(default_factory=list)


@dataclass
class Draft:
    rows: list[DraftRow]
    notes: list[str]


class ExtractionError(Exception):
    pass


def is_configured() -> bool:
    return bool(os.environ.get("ANTHROPIC_API_KEY") or os.environ.get("ANTHROPIC_AUTH_TOKEN"))


def prepare_image(data: bytes) -> tuple[bytes, str]:
    """Fix phone rotation, downscale, and re-encode as JPEG. Returns (bytes, media_type)."""
    try:
        img = Image.open(BytesIO(data))
        img = ImageOps.exif_transpose(img)
    except Exception as exc:  # noqa: BLE001
        raise ExtractionError("Could not open the photo. Please upload a JPG or PNG image.") from exc
    img = img.convert("RGB")
    img.thumbnail((MAX_SIDE, MAX_SIDE))
    out = BytesIO()
    img.save(out, format="JPEG", quality=90)
    return out.getvalue(), "image/jpeg"


def canonical_rider(name: str, riders: list[str]) -> tuple[str, list[str]]:
    by_key = {r.lower().replace("-", "").replace(" ", ""): r for r in riders}
    key = name.strip().lower().replace("-", "").replace(" ", "")
    if key in by_key:
        return by_key[key], []
    close = difflib.get_close_matches(key, list(by_key), n=1, cutoff=0.7)
    if close:
        return by_key[close[0]], [f'rider read as "{name}"']
    return name.strip(), [f'unknown rider "{name}"']


def read_board(
    image: bytes,
    media_type: str,
    riders: list[str],
    clients: list[Client],
    model: str,
) -> BoardReading:
    import anthropic  # imported lazily so the app runs without the SDK for manual entry

    client_names = ", ".join(c.name for c in clients) or "(no client list yet)"
    user_text = (
        f"Known riders: {', '.join(riders)}\n"
        f"Known clients: {client_names}\n\n"
        "Transcribe every delivery entry on this board."
    )
    api = anthropic.Anthropic()
    try:
        response = api.beta.messages.parse(
            model=model,
            max_tokens=16000,
            system=SYSTEM_PROMPT,
            output_config={"effort": "high"},
            # If the request is ever declined by a safety classifier, let the API
            # retry it on a suitable fallback model inside the same call.
            betas=["server-side-fallback-2026-07-01"],
            fallbacks="default",
            messages=[{
                "role": "user",
                "content": [
                    {
                        "type": "image",
                        "source": {
                            "type": "base64",
                            "media_type": media_type,
                            "data": base64.standard_b64encode(image).decode("ascii"),
                        },
                    },
                    {"type": "text", "text": user_text},
                ],
            }],
            output_format=BoardReading,
        )
    except anthropic.AuthenticationError as exc:
        raise ExtractionError("The Anthropic API key is invalid. Check ANTHROPIC_API_KEY.") from exc
    except anthropic.RateLimitError as exc:
        raise ExtractionError("The photo reader is busy (rate limited). Try again in a minute.") from exc
    except anthropic.APIStatusError as exc:
        raise ExtractionError(f"The photo reader returned an error ({exc.status_code}).") from exc
    except anthropic.APIConnectionError as exc:
        raise ExtractionError("Could not reach the photo reader. Check the internet connection.") from exc

    if response.stop_reason == "refusal":
        raise ExtractionError("The photo reader declined this image. Enter the rows by hand.")
    if response.stop_reason == "max_tokens" or response.parsed_output is None:
        raise ExtractionError("The photo reader could not finish reading the board. Enter the rows by hand.")
    return response.parsed_output


def to_draft(
    reading: BoardReading,
    riders: list[str],
    clients: list[Client],
    aliases: dict[str, str],
) -> Draft:
    """Normalize names, merge duplicate client lines per rider, attach review flags."""
    merged: dict[tuple[str, str], DraftRow] = {}
    for e in reading.entries:
        if e.quantity <= 0:
            continue
        rider, rider_flags = canonical_rider(e.rider, riders)
        client, client_flags = match_client(e.client, clients, aliases)
        flags = rider_flags + client_flags
        if e.unclear:
            flags.append("unclear handwriting" + (f": {e.note}" if e.note else ""))
        key = (rider, client)
        if key in merged:
            merged[key].qty += e.quantity
            merged[key].flags.append("combined duplicate entries")
            merged[key].flags.extend(f for f in flags if f not in merged[key].flags)
        else:
            merged[key] = DraftRow(rider, client, e.quantity, flags)

    order = {r: i for i, r in enumerate(riders)}
    rows = sorted(merged.values(), key=lambda r: order.get(r.rider, len(order)))

    notes = list(reading.notes)
    if reading.riders_off:
        notes.append("No sales / OFF: " + ", ".join(reading.riders_off))
    for item in reading.left_out:
        notes.append(f'Left out (not a "quantity client" entry): {item}')
    return Draft(rows, notes)
