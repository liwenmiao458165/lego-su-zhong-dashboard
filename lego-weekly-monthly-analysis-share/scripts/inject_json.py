#!/usr/bin/env python3
"""Inject JSON data into HTML files by replacing the embedded const DATA = {...} block."""
import json
import re
import sys

def inject_json(html_path, json_path, output_path=None):
    if output_path is None:
        output_path = html_path

    with open(json_path, 'r', encoding='utf-8') as f:
        data = json.load(f)

    with open(html_path, 'r', encoding='utf-8') as f:
        html = f.read()

    # Find the const DATA = { marker
    marker = 'const DATA = '
    idx = html.find(marker)
    if idx == -1:
        print(f"ERROR: marker '{marker}' not found in {html_path}")
        return False

    start = idx + len(marker)

    # Find the end of the JSON object by brace matching
    depth = 0
    in_string = False
    escape = False
    end = start
    for i in range(start, len(html)):
        c = html[i]
        if escape:
            escape = False
            continue
        if c == '\\':
            escape = True
            continue
        if c == '"' and not escape:
            in_string = not in_string
            continue
        if in_string:
            continue
        if c == '{':
            depth += 1
        elif c == '}':
            depth -= 1
            if depth == 0:
                end = i + 1
                break

    # Build new JSON string
    json_str = json.dumps(data, ensure_ascii=False, separators=(',', ':'))

    # Replace
    new_html = html[:start] + json_str + html[end:]

    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(new_html)

    print(f"OK: {html_path}")
    print(f"  JSON: {len(json_str)} chars")
    print(f"  HTML: {len(new_html)} chars")
    return True

if __name__ == '__main__':
    base = '/Users/a123/WorkBuddy/Claw/outputs'

    inject_json(
        f'{base}/weekly_analysis.html',
        f'{base}/weekly_analysis.json'
    )

    inject_json(
        f'{base}/monthly_analysis.html',
        f'{base}/monthly_analysis.json'
    )

    print("\nDone!")
