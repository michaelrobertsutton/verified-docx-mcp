"""Explicit, digest-bound formatting for live text replacements."""
from __future__ import annotations

import re
from typing import Any
from urllib.parse import urlsplit

from .errors import ErrorCode, _make_error

BOOLEANS = ('bold', 'italic', 'strikeThrough', 'doubleStrikeThrough', 'subscript', 'superscript')
FIELDS = {*BOOLEANS, 'underline', 'color', 'name', 'size'}


def validate_segments(segments: list[dict[str, Any]] | None, replacement: str):
    if segments is None:
        return None
    def bad(message):
        raise _make_error(ErrorCode.INVALID_INPUT, message)
    if not isinstance(segments, list) or len(segments) > 500:
        bad('replacement_segments must be a list of at most 500 segments')
    result = []
    for segment in segments:
        if not isinstance(segment, dict) or set(segment) - {'text', 'format', 'hyperlink'} or not {'text', 'format'} <= set(segment):
            bad('Each segment requires text and format, with optional hyperlink')
        text, font = segment['text'], segment['format']
        if not isinstance(text, str) or not text or any(ord(c) < 32 and c not in '\t\v' for c in text):
            bad('Segment text must be nonempty inline text; use vertical tab for a soft line break')
        if not isinstance(font, dict) or set(font) != FIELDS:
            bad('Each segment format must explicitly specify all font fields: ' + ', '.join(sorted(FIELDS)))
        if any(type(font[k]) is not bool for k in BOOLEANS):
            bad('Segment font flags must be boolean')
        if font['subscript'] and font['superscript']:
            bad('A segment cannot be both subscript and superscript')
        if font['underline'] not in {'None', 'Single', 'Double'}:
            bad('Segment underline must be None, Single, or Double')
        if not isinstance(font['color'], str) or not re.fullmatch(r'#[0-9a-fA-F]{6}', font['color']):
            bad('Segment color must be #RRGGBB')
        if not isinstance(font['name'], str) or not font['name'].strip() or len(font['name']) > 128:
            bad('Segment font name must be nonempty')
        if type(font['size']) not in (int, float) or not 1 <= font['size'] <= 1638 or font['size'] * 2 != int(font['size'] * 2):
            bad('Segment font size must be 1..1638 in half-point increments')
        link = segment.get('hyperlink')
        if link is not None and (not isinstance(link, str) or urlsplit(link).scheme not in {'https', 'http', 'mailto'} or any(ord(c) < 32 for c in link)):
            bad('Segment hyperlink must be an explicit http, https, or mailto URL, or null')
        result.append({'text': text, 'format': {**font, 'color': font['color'].upper()}, 'hyperlink': link})
    if ''.join(s['text'] for s in result) != replacement:
        bad('Segment text must concatenate exactly to replace')
    return result
