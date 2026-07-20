"""
data_extraction.ai_utils.entity_chains
======================================

The predefined **"Specific entities"** column of the AI Review column
analysis: the extraction prompt, the parser for the chain format it asks
the LLM to produce, and the Excel writer that fans parsed chains out into
a component sheet.

Chain format (one per relationship, ``;``-separated)::

    Reference|System Info|Process|Personal|QuantityValue

``Reference`` is mandatory; missing optional components are ``#``.
Components are joined with a pipe (``|``). The pipe is used instead of a
hyphen because References routinely contain hyphens themselves
(``FAA AC 120-76D``, ``RTCA DO-178C``) — a hyphen connector made the field
boundaries ambiguous, whereas ``|`` never occurs inside a reference or
value, so :func:`parse_chain` is a simple positional split. Chains without
a Reference are dropped, per the prompt's own rule. Legacy hyphen-format
values (produced before this change) are still parsed via a right-split
fallback.

Used in two places:

* automatically during a column-analysis run — when the "Specific
  entities" column is among the analyzed columns, every saved row's
  chains are re-parsed into an ``Entities <run sheet>`` sheet of the
  storage workbook (one row per chain, one column per component);
* standalone from the *Unique elements* tab via
  :func:`entities_from_workbook` — parse any column of any Excel sheet
  that holds such chain values.
"""

from __future__ import annotations

import os

ENTITY_COLUMN = "Specific entities"

COMPONENTS = ("Reference", "System Info", "Process", "Personal", "QuantityValue")

# Component connector. A pipe is used (not a hyphen) because references
# routinely contain hyphens (FAA AC 120-76D, RTCA DO-178C), which made a
# hyphen connector ambiguous to split. The pipe never appears inside a
# reference or value, so parsing is a plain positional split.
COMPONENT_SEP = "|"

ENTITY_PROMPT = """analyze aviation texts and extract relationships into a strict format based on five specific categories:

1. Reference (Mandatory): Any regulatory, legal, national/international authority rules, or technical industry standards. (e.g., EASA AMC1 ORO.GEN.200, FAA FAR Part 21, ICAO Annex 19, RTCA DO-178C, ISO 9001).
2. System Info (Optional): Aircraft models, engine types, components, parts, systems, software, or ground tools. (e.g., landing gear)
3. Process (Optional): Technical or operational actions like inspection, overhaul, auditing, software verification, or risk assessment. (e.g., defect rectification)
4. Personal (Optional): Accountable managers, certifying staff, pilots, operators, manufacturers, airlines, or competent authorities. (e.g., CAMO, FAA inspector)
5. QuantityValue (Optional): Physical measurements, numerical limits, intervals, or tolerances including their units. (e.g., 500 hours, 25 kg, 15°C)

### Formatting Rules:
- Format every extracted chain exactly as: Reference|System Info|Process|Personal|QuantityValue
- Use the pipe character '|' to separate the five components. Do NOT use hyphens as separators (references such as FAA AC 120-76D contain hyphens themselves).
- The "Reference" component must ALWAYS exist. If no Reference is found, do not create a chain.
- The order of components must be strictly maintained. Do not include spaces around the pipes. Never use the '|' character inside a component value.
- If an optional component (System, Process, Personal, or QuantityValue) is missing, replace it with the '#' character.
- If multiple distinct chains or relationships are found in the text, separate each chain with a semicolon (;).
- Return ONLY the final formatted string. No explanations, no introductory text, no markdown code blocks.

### Examples for Training:
- Input: "According to FAA AC 120-76D, the operator must evaluate the electronic flight bag software every 6 months."
  Output: FAA AC 120-76D|electronic flight bag software|evaluation|operator|6 months

- Input: "RTCA DO-178C states that the software development team shall perform code reviews for Level A software."
  Output: RTCA DO-178C|Level A software|code reviews|software development team|#

- Input: "Per ISO 9001:2015, the internal auditor must verify the calibration of all torque wrenches weighing over 5 kg."
  Output: ISO 9001:2015|torque wrenches|calibration verification|internal auditor|over 5 kg"""


def is_entity_column(name):
    """True when a column name is the Specific-entities column (any case)."""
    return str(name or "").strip().lower() == ENTITY_COLUMN.lower()


def clean_chain_text(text):
    """Strip a markdown code fence (```…```) and surrounding whitespace from
    a raw LLM reply, leaving the bare chain string."""
    raw = str(text if text is not None else "").strip()
    if raw.startswith("```"):
        lines = raw.splitlines()
        lines = lines[1:]                       # drop the ``` / ```text line
        while lines and lines[-1].strip().startswith("```"):
            lines = lines[:-1]
        raw = "\n".join(lines).strip()
    return raw


def looks_like_chain(text):
    """True when a raw LLM reply *is* entity-chain text rather than JSON or
    prose, so it can be recorded as-is instead of re-prompting for JSON.

    Deliberately strict: every ``;``-separated segment must use the pipe
    connector and at least one must parse into a valid chain. Requiring the
    pipe is what keeps ordinary prose (which is full of hyphens) from being
    mistaken for a chain; a Reference-less segment may still fail to parse
    and is simply dropped later, so only one valid chain is required. A
    JSON-looking reply is rejected outright so this stays correct even when
    called before a JSON parse."""
    raw = clean_chain_text(text)
    if not raw or COMPONENT_SEP not in raw:
        return False
    if raw.startswith("{") or raw.startswith("["):
        return False        # JSON (which may well contain pipes in values)
    segments = [s.strip() for s in raw.split(";") if s.strip()]
    if not segments or not all(COMPONENT_SEP in s for s in segments):
        return False
    return any(parse_chain(s) is not None for s in segments)


def parse_chain(chain):
    """Parse one ``Reference|System|Process|Personal|Quantity`` chain into a
    ``{component: value}`` dict, or ``None`` for an empty / Reference-less
    chain. ``#`` placeholders become "". Components are pipe-separated, so
    this is a plain positional split; the first field is the (mandatory)
    Reference and the next four are the optional components. Legacy
    hyphen-format chains (no pipe present) fall back to the old right-split
    that tolerates hyphens inside the Reference."""
    raw = str(chain or "")
    if COMPONENT_SEP not in raw:
        return _parse_chain_legacy_hyphen(raw)
    parts = [("" if p.strip() == "#" else p.strip())
             for p in raw.split(COMPONENT_SEP)]
    if not any(parts):
        return None
    ref = parts[0]
    if not ref:
        return None  # the prompt forbids Reference-less chains
    # Positional: 4 optional components after the Reference; anything beyond
    # the fifth field (a stray '|' in a value) is appended to QuantityValue.
    tail = parts[1:5] + [""] * (4 - len(parts[1:5]))
    if len(parts) > 5:
        extra = COMPONENT_SEP.join(p for p in parts[5:] if p)
        tail[3] = (tail[3] + COMPONENT_SEP + extra).strip(COMPONENT_SEP) \
            if extra else tail[3]
    return dict(zip(COMPONENTS, [ref] + tail))


def _parse_chain_legacy_hyphen(chain):
    """Parse a pre-pipe ``Reference-System-Process-Personal-Quantity`` chain.
    References may contain hyphens, so the last four ``-``-fields are the
    optional components and everything before them joins back into the
    Reference. Kept so entity sheets/values produced before the pipe switch
    still parse."""
    parts = [p.strip() for p in str(chain or "").split("-")]
    parts = [("" if p == "#" else p) for p in parts]
    if not any(parts):
        return None
    if len(parts) >= 5:
        ref = "-".join(str(chain).split("-")[:len(parts) - 4]).strip()
        ref = "" if ref == "#" else ref
        tail = parts[-4:]
    else:  # malformed / short chain: first field is the Reference
        ref = parts[0]
        tail = parts[1:] + [""] * (4 - len(parts[1:]))
    if not ref:
        return None  # the prompt forbids Reference-less chains
    return dict(zip(COMPONENTS, [ref] + tail))


def parse_chains(cell):
    """Parse every ``;``-separated chain of one cell; skips empty and
    Reference-less chains. Returns a list of component dicts, each with the
    raw chain kept under ``"Chain"``."""
    text = str(cell if cell is not None else "")
    if not text.strip() or cell != cell:  # None / NaN / blank
        return []
    out = []
    for chain in text.split(";"):
        chain = chain.strip()
        if not chain:
            continue
        parsed = parse_chain(chain)
        if parsed is not None:
            parsed["Chain"] = chain
            out.append(parsed)
    return out


def extract_rows(records, column=ENTITY_COLUMN, section_key="Section"):
    """Parse the entity column of analysis-row dicts into sheet rows:
    ``Section | Reference | System Info | Process | Personal |
    QuantityValue | Chain`` — one row per chain."""
    rows = []
    for rec in records:
        # the column may be typed with different capitalisation
        cell = next((rec[k] for k in rec if is_entity_column(k)), None) \
            if is_entity_column(column) else rec.get(column)
        for parsed in parse_chains(cell):
            rows.append({section_key: rec.get(section_key, ""), **parsed})
    return rows


def write_entity_sheet(path, rows, sheet_name, section_key="Section"):
    """(Re)write one component sheet from parsed chain rows (idempotent —
    an existing sheet of that name is replaced). Returns the sheet name."""
    import pandas as pd

    cols = [section_key, *COMPONENTS, "Chain"]
    df = pd.DataFrame(rows, columns=cols)
    sheet_name = str(sheet_name)[:31]
    if os.path.exists(path):
        with pd.ExcelWriter(path, engine="openpyxl", mode="a",
                            if_sheet_exists="replace") as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    else:
        with pd.ExcelWriter(path, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    return sheet_name


def entities_from_workbook(path, sheet, column, out_sheet=None):
    """Standalone: parse the chain values of ``column`` in ``sheet`` of an
    Excel file into an ``Entities <sheet>`` component sheet of the same
    file. Returns a summary dict {sheet, rows, chains}."""
    import pandas as pd

    df = pd.read_excel(path, sheet_name=sheet)
    if column not in df.columns:
        raise ValueError(f"Column '{column}' not found in sheet '{sheet}'.")
    section_col = "Section" if "Section" in df.columns else None
    rows = []
    for rec in df.to_dict("records"):
        for parsed in parse_chains(rec.get(column)):
            rows.append({"Section": rec.get(section_col, "") if section_col
                         else "", **parsed})
    out = write_entity_sheet(path, rows, out_sheet or f"Entities {sheet}"[:31])
    return {"sheet": out, "rows": len(df), "chains": len(rows)}
