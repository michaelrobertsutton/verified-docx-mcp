"""Explicit, digest-bound formatting for live text replacements.

Issue #108: a segment's ``format`` may name only the fields that differ; every field it
leaves out is inherited from the effective font of the FIRST replaced run (the pane's
read-only ``replacement_font`` op). The server resolves partial segments to full ones before
anything is written -- the pane only ever receives complete segments -- and sends the
baseline it used as ``expected_first_font`` so the pane refuses (``FORMAT_DRIFT``) if that
run's formatting changed before the write.
"""
from __future__ import annotations

import re
from typing import Any
from urllib.parse import urlsplit

from .errors import ErrorCode, _make_error

BOOLEANS = ('bold', 'italic', 'strikeThrough', 'doubleStrikeThrough', 'subscript', 'superscript')
FIELDS = {*BOOLEANS, 'underline', 'color', 'name', 'size'}
PARTIAL_CAPABILITY = 'partial_font_segments_v1'


def _bad(message):
    raise _make_error(ErrorCode.INVALID_INPUT, message)


def _field_error(key: str, value: Any) -> str | None:
    """Why *value* is not a valid ``key`` font field, or None when it is."""
    if key in BOOLEANS:
        return None if type(value) is bool else 'Segment font flags must be boolean'
    if key == 'underline':
        return None if value in {'None', 'Single', 'Double'} else 'Segment underline must be None, Single, or Double'
    if key == 'color':
        return None if isinstance(value, str) and re.fullmatch(r'#[0-9a-fA-F]{6}', value) else 'Segment color must be #RRGGBB'
    if key == 'name':
        return None if isinstance(value, str) and value.strip() and len(value) <= 128 else 'Segment font name must be nonempty'
    if key == 'size':
        size_ok = type(value) in (int, float) and 1 <= value <= 1638 and value * 2 == int(value * 2)
        return None if size_ok else 'Segment font size must be 1..1638 in half-point increments'
    return f'Unknown font field {key!r}'


def validate_segments(segments: list[dict[str, Any]] | None, replacement: str, *, partial: bool = False):
    """Validate and normalize ``replacement_segments``.

    ``partial=False`` (the default, and what every pane readback uses) requires all ten
    font fields per segment. ``partial=True`` accepts any SUBSET of the fields: unknown keys
    are still refused and every field that IS present obeys the same type/value rules.
    """
    if segments is None:
        return None
    if not isinstance(segments, list) or len(segments) > 500:
        _bad('replacement_segments must be a list of at most 500 segments')
    result = []
    for segment in segments:
        if not isinstance(segment, dict) or set(segment) - {'text', 'format', 'hyperlink'} or not {'text', 'format'} <= set(segment):
            _bad('Each segment requires text and format, with optional hyperlink')
        text, font = segment['text'], segment['format']
        if not isinstance(text, str) or not text or any(ord(c) < 32 and c not in '\t\v' for c in text):
            _bad('Segment text must be nonempty inline text; use vertical tab for a soft line break')
        if not isinstance(font, dict):
            _bad('Each segment format must be an object of font fields')
        unknown = set(font) - FIELDS
        if unknown:
            _bad(f'Unknown segment format field(s) {sorted(unknown)}; allowed: ' + ', '.join(sorted(FIELDS)))
        if not partial and set(font) != FIELDS:
            _bad('Each segment format must explicitly specify all font fields: ' + ', '.join(sorted(FIELDS)))
        for key in sorted(font):
            problem = _field_error(key, font[key])
            if problem:
                _bad(problem)
        if font.get('subscript') and font.get('superscript'):
            _bad('A segment cannot be both subscript and superscript')
        link = segment.get('hyperlink')
        if link is not None and (not isinstance(link, str) or urlsplit(link).scheme not in {'https', 'http', 'mailto'} or any(ord(c) < 32 for c in link)):
            _bad('Segment hyperlink must be an explicit http, https, or mailto URL, or null')
        normalized = dict(font)
        if 'color' in normalized:
            normalized['color'] = normalized['color'].upper()
        result.append({'text': text, 'format': normalized, 'hyperlink': link})
    if ''.join(s['text'] for s in result) != replacement:
        _bad('Segment text must concatenate exactly to replace')
    return result


def is_partial(segments: list[dict[str, Any]] | None) -> bool:
    """True when any segment leaves a font field to be inherited."""
    return any(set(s['format']) != FIELDS for s in segments or [])


def normalize_baseline(font: Any) -> dict[str, Any]:
    """The pane's reported first-run font, as a baseline: the ten fields, colour upper-cased.
    A field the pane could not report (Word answers null/"" for a mixed range) stays as
    reported; resolution refuses to inherit it."""
    if not isinstance(font, dict) or set(font) < FIELDS:
        raise _make_error(ErrorCode.LIVE_OP_FAILED, 'The pane returned a malformed replacement_font reply',
                          {'stage': 'replacement_font_reply'})
    out = {k: font[k] for k in sorted(FIELDS)}
    if isinstance(out['color'], str):
        out['color'] = out['color'].upper()
    return out


def resolve_segments(segments: list[dict[str, Any]], baseline: dict[str, Any], replacement: str):
    """Fill every field a segment leaves out from *baseline* (the effective font of the first
    replaced run) and return complete, validated segments. An inherited field whose baseline
    value cannot be written back explicitly (Word reports a value outside the supported
    schema) is refused with INVALID_INPUT naming the field."""
    resolved = []
    for segment in segments:
        font = dict(segment['format'])
        for key in sorted(FIELDS - set(font)):
            if _field_error(key, baseline.get(key)):
                _bad(f'The first replaced run\'s {key} ({baseline.get(key)!r}) cannot be inherited; '
                     f'specify {key} explicitly in the segment format')
            font[key] = baseline[key]
        resolved.append({**segment, 'format': font})
    return validate_segments(resolved, replacement)


def request_replacement_font(session, payload: dict[str, Any]) -> dict[str, Any]:
    """Ask the pane (read-only op ``replacement_font``) for the effective font of the first
    run of the target the replace payload names. *session* must be a non-observing view so
    the lookup never re-arms the caller's write baseline."""
    from .live import write_mode as wm
    from .live.session import LiveDisconnected, LiveOpFailed

    try:
        reply = session.request_threadsafe('replacement_font', payload)
    except LiveOpFailed as exc:
        raise _make_error(wm.classify_op_failed(exc), exc.message, {'pane_code': exc.code}) from exc
    except LiveDisconnected as exc:
        raise _make_error(ErrorCode.LIVE_DISCONNECTED, str(exc)) from exc
    return normalize_baseline(reply.get('font') if isinstance(reply, dict) else None)
