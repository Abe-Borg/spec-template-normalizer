from paragraph_rules import compute_skip_reason, is_copyright_notice, is_specifier_note


def test_is_copyright_notice_cases() -> None:
    cases = [
        ("Copyright 2011 by The American Institute of Architects (AIA)", True),
        ("Copyright 2019 by The American Institute of Architects (AIA)", True),
        (
            "Exclusively published and distributed by Architectural Computer Services, Inc. (ARCOM) for the AIA",
            True,
        ),
        ("Exclusively published and distributed by XYZ Corp.", True),
        ("Comply with copyright regulations.", False),
        ("Copyright protection applies to all documents.", False),
        ("", False),
        ("SECTION 22 05 29", False),
    ]

    for raw_text, expected in cases:
        assert is_copyright_notice(raw_text) is expected


def test_is_specifier_note_cases() -> None:
    cases = [
        ("Retain first paragraph below if Contractor is required to assume responsibility for design.", True),
        ("Retain both paragraphs below if shop or field welding is required.", True),
        ("Retain Sections in subparagraphs below that contain requirements...", True),
        ("Retain definitions remaining after this Section has been edited.", True),
        ("Retain option in first paragraph below...", True),
        ("Revise this Section by deleting and inserting text...", True),
        ("Verify that Section titles referenced in this Section are correct...", True),
        ("See Editing Instruction No. 1 in the Evaluations...", True),
        ("Specify parts in first three subparagraphs below...", True),
        ("Paragraph below is defined in Section 013300...", True),
        ("Trapeze pipe hanger in paragraph below requires calculating and detailing at each use.", True),
        (
            "Metal framing system in first paragraph below requires calculating and detailing at each use.",
            True,
        ),
        (
            "Retain first paragraph below if Section 099113 is in Project Manual. Revise reference...",
            True,
        ),
        ("Option:  Thermal-hanger shield inserts may be used.", True),
        ("Manufacturers' catalogs indicate that copper pipe hangers are small...", True),
        ("High-compressive-strength inserts may permit use of shorter shields...", True),
        # False-positive guards
        ("Retain existing pipe supports where indicated.", False),
        ("Verify all dimensions in the field before fabrication.", False),
        ("Option: Provide seismic bracing where required by code.", False),
        ("Section Includes:", False),
        ("Adjustable Steel Clevis Hangers: (MSS Type 1.) B-Line B 3100", False),
        (
            "Install hangers and supports to allow controlled thermal and seismic movement of piping systems.",
            False,
        ),
        (
            "Structural Steel Welding Qualifications: Qualify procedures and personnel according to AWS D1.1",
            False,
        ),
        ("", False),
    ]

    for raw_text, expected in cases:
        assert is_specifier_note(raw_text) is expected


def test_compute_skip_reason_integration() -> None:
    cases = [
        (
            "Copyright 2011 by The American Institute of Architects (AIA)",
            False,
            False,
            "copyright_notice",
        ),
        ("Retain first paragraph below...", False, False, "specifier_note"),
        ("Retain first paragraph below...", True, False, "sectPr"),
        ("Retain first paragraph below...", False, True, "in_table"),
        ("Install hangers and supports...", False, False, None),
    ]

    for raw_text, contains_sectpr, in_table, expected in cases:
        assert compute_skip_reason(raw_text, contains_sectpr, in_table) == expected
