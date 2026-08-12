---
name: sound-noun
description: Extract the main sound-event noun or concise noun phrase from caption-annotation spreadsheets or text and output timestamped, track-numbered, track-named TXT lines. Use for Sound Noun project files containing caption numbers, track numbers, track names/descriptions, start/end times, and annotation text.
---

# Sound Noun

## Purpose

Convert audio-caption annotation data into a compact TXT-style list containing:

`Start time<TAB>End time<TAB>Track Number – Sound event noun<TAB>Track`

Example:

Input:

`3    2    Breathing    Brief faint breathing.    1.400    1.767    Not visible / offscreen source    Slight    Brief faint breathing.`

Output:

`1.400    1.767    2 – Breathing    Breathing`

The fourth field is the row's `Track` value, copied verbatim. It names the label
track this line lands on when the file is imported into Audacity, which is why it
is the raw column value and not the sound noun.

## Core Task

For every caption row:

1. Preserve the exact start time.
2. Preserve the exact end time.
3. Preserve the assigned track number.
4. Identify the principal noun naming the audible sound event.
5. Copy the `Track` value verbatim as the fourth field.
6. Output one line per caption in the same chronological/row order as the source.

## Output Format

Use tab-separated fields:

`StartTime<TAB>EndTime<TAB>TrackNumber – Sound event noun<TAB>Track`

Example:

```text
1.400	1.767	2 – Breathing	Breathing
1.733	30.033	3 – Humming	Ambient hum
4.033	6.700	1 – Male speech	Male Speech
```

The fourth field may be left off entirely, giving the older three-field shape,
which still imports — the label tracks are then named by number alone. Include it
whenever the source has a `Track` column.

Do not include:
- caption numbers
- track descriptions
- source visibility
- source descriptions
- prominence
- explanatory commentary
- column headers, unless the user explicitly requests them

When exporting a file, save as UTF-8 plain text (`.txt`).

## Sound-Event Noun Selection

Choose the shortest natural noun or noun phrase that accurately identifies the acoustic event.

Prefer:
- `Breathing`
- `Humming`
- `Footsteps`
- `Clacking`
- `Rustling`
- `Scraping`
- `Bird chirping`
- `Male speech`
- `Female laughter`
- `Engine rumble`
- `Metal movement`
- `Wailing siren`

Avoid unnecessary acoustic modifiers such as:
- faint
- loud
- brief
- continuous
- distant
- high-pitched
- low-frequency
- rhythmic
- resonant
- muffled
- prominent

Retain a modifier only when it is necessary to identify the event itself, such as:
- `Male speech`
- `Female speech`
- `Bird chirping`
- `Engine rumble`
- `Door creak`
- `Metal clanking`

Usually keep the sound-event label to one to three words.

## Evidence Priority

Determine the sound noun using all available fields, with this practical priority:

1. `Track`
2. `AnnotationText`
3. `TrackDescription`

Use the most semantically accurate event name rather than mechanically copying a field.

If the track name is already a clean sound-event noun or noun phrase, normally use it.

If the track name is vague, overly broad, awkward, or contains multiple phenomena, use the annotation text and track description to choose the principal audible event.

Examples:

- Track: `Various Bird Noises`
  Annotation: `various birds chirping wildly`
  → `Bird chirping`

- Track: `Chirp-friction`
  Annotation: `faint closely spaced sounds with a lightly chirping quality and subtle friction-like texture`
  → choose the principal event supported by the caption, usually `Chirping` unless friction is clearly the actual event.

- Track: `Metal Movement`
  Annotation: `metal scraping and shifting`
  → use the most specific principal event supported by the annotation, such as `Metal scraping`, rather than preserving an unnecessarily vague label.

## Multiple Sounds in One Caption

When a caption contains several sound descriptions:

- Select the main or dominant sound event represented by that caption/track.
- Do not create multiple output lines unless the source itself contains separate caption rows.
- If a compact compound noun phrase is necessary to preserve the event identity, keep it concise.

Example:

`slow, spaced-out synth bass track with melody pads and metallic textural hits`
→ `Synth bass` if the bass is the principal event.

## Speech and Vocalizations

Use concise event labels such as:
- `Male speech`
- `Female speech`
- `Speech`
- `Laughter`
- `Female laughter`
- `Crying`
- `Screaming`
- `Grunting`
- `Breathing`
- `Whispering`

Do not include language, speaking rate, emotion, pitch, distance, or volume unless needed to distinguish the event requested by the user.

## Workbook Handling

When processing spreadsheets:

1. Read all relevant caption rows.
2. Preserve each row's exact `StartTime(s)`, `EndTime(s)`, `TrackNumber` and `Track`.
3. Use the `Track`, `TrackDescription`, and `AnnotationText` fields to determine the noun.
4. A workbook usually holds several sheets carrying this schema — `Completion` and `Refinement` alongside sheets that do not, such as `Summary`. Find the sheets carrying all ten headers, read that one silently when exactly one does, and otherwise ask which to use rather than picking one.
5. Do not omit repeated captions merely because the sound noun or track number repeats.
6. Preserve the source row order unless the user asks for sorting.
7. Do not renumber tracks.
8. These sheets are typically sorted alphabetically by `Track`, so neither `TrackNumber` nor time runs in order down the rows, and rows of one track are not necessarily adjacent. Read `TrackNumber` and `Track` off each row itself; never carry a value down from the row above. Output order does not matter to the import, which groups by number and re-sorts by time.
9. `TrackNumber` and `Track` pair up one-to-one. Check it: if one number appears with two different names, say so and stop rather than choosing one. If one name appears under two numbers, mention it — the import still works, but the workbook has one track split in two.

## Precision Rules

- Keep timestamp precision exactly as supplied whenever possible.
- Do not round `1.400` to `1.4`.
- Do not convert seconds to timecode unless requested.
- Use an en dash with spaces between track number and noun: `2 – Breathing`.
- Use tabs between all four output components, not spaces.
- Copy the fourth field from `Track` verbatim — same wording, same capitalization,
  including a trailing period if the column has one. Trim leading and trailing
  whitespace and nothing else. Do not substitute the sound noun for it.
- The fourth field must not contain a tab. A `Track` value holding one is a data
  problem worth reporting rather than silently repairing.
- Every row sharing a `TrackNumber` must carry the same fourth field. The
  importer refuses a file where one number carries two different names.

## Quality Checks

Before delivering the result, verify:

- Every source caption row has one output line.
- Every line includes a start time.
- Every line includes an end time.
- Every line includes the correct track number.
- Every line includes a sound-event noun or concise noun phrase.
- Every line includes the `Track` value as a fourth field, or no line does.
- The same track number never carries two different fourth fields, in any
  spelling or capitalization — this is the one thing that stops the import.
- No caption number has accidentally replaced the track number.
- No descriptive sentence has been copied in place of the noun.
- No sound noun has been copied in place of the `Track` value, or the reverse.
- Repeated track numbers remain repeated where the source requires them.
- Output order matches the source.

## File Naming

If the user supplies one source file and asks for an exported TXT without specifying a name, use the source filename stem plus `.txt`.

Examples:

- `Twirp.xlsx` → `Twirp.txt`
- `Rack R0.xlsx` → `Rack R0.txt`

If multiple source files are supplied, create one TXT per source file unless the user requests a combined output.

## Response Behavior

For file-processing requests:
- Produce the requested TXT file(s).
- Provide direct download links.
- Keep the chat response minimal.
- Do not add an explanatory table unless requested.

For pasted caption data:
- Return the TXT-formatted lines directly in a code block for easy copy/paste.
