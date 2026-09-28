# DocxFill prefill example

#

# Load this via "Load Prefill..." in the app. Lines outside a

# {{field}} / \[option] block (like this comment header) are ignored,

# so you can annotate the file freely.

#

# Fields below correspond to a typical civil engineering fee proposal

# template. Adjust the field names to match whatever {{fields}} your

# own template actually uses - they must match exactly, including

# capitalization and spacing.

{{AHJ}}
\[City of Mukilteo]
\[City of Everett]
\[Snohomish County]
\[City of Marysville]
\[City of Lake Stevens]
\[City of Seattle]
\[City of Kent]
\[King County]

{{STORM MANUAL}}
\[2024 Snohomish County Drainage Manual]
\[2024 Stormwater Management Manual for Western Washington]
\[2024 Stormwater Management Manual for Eastern Washington]
\[2026 King County Surface Water Design Manual]
\[2026 Seattle Stormwater Code and Manual]


{{Sewer District}}
\[Alderwood Water & Wastewater District]
\[City of Mukilteo Public Works]
\[City of Everett Public Works]
\[Silver Lake Water & Sewer District]
\[Seattle Public Utilities]

{{Water District}}
\[Alderwood Water & Wastewater District]
\[City of Everett Public Works]
\[Silver Lake Water & Sewer District]
\[Seattle Public Utilities]

{{ZONE DESCRIPTION}}
\[Single Family Residential]
\[Multi-Family Residential]
\[General Commercial]
\[Light Industrial]

{{PROJECT DESCRIPTION}}
\[subdivide the parcel into single-family residential lots]
\[construct a multi-family residential development]
\[construct a commercial building and associated site improvements]

# This option demonstrates a nested field: choosing it inserts

# {{PROJECT NAME}} into the document, which is a REAL field used

# elsewhere in the template (e.g. a "Project: {{PROJECT NAME}}"

# header line), so it just reuses whatever value you enter for

# PROJECT NAME - no separate entry needed for it.

#

# Note: pick nesting targets like this deliberately. A field that

# already sits right next to {{SPECIFY}} in the template's own

# sentence (e.g. this template separately reads "...{{SPECIFY}} dated

# {{ARCH DATE}}.") would end up repeated twice in the same sentence if

# also nested here - reusing a field from an unrelated part of the

# document, like PROJECT NAME, avoids that.

{{SPECIFY}}
\[the project architect for {{PROJECT NAME}}]
\[the project surveyor, per the boundary survey provided]
\[the owner's representative]

# This option demonstrates a nested field that is NOT used anywhere

# else in the template. Picking it will add "ARCHITECT FIRM NAME" as

# its own entry under "Additional fields" in the side panel, since it

# only exists inside this prefill text.

{{AHJ DEVELOPMENT STANDARD}}
\[2026 Seattle Construction Standards(Specifications and Plans)]
\[per {{ARCHITECT FIRM NAME}}'s site plan and applicable municipal code]
\[per the applicable municipal code and current adopted standards]
