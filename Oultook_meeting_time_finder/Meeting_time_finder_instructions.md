Do not simplify the task. Do not provide pseudocode. Do not omit file export. I need full working VBA code, not a conceptual example.

Before outputting the code, think through the slot intersection logic carefully so that the script does not falsely mark blocked time as free.

Write a complete, production-ready VBA macro for Microsoft Outlook that finds common meeting slots across multiple attendees by reading each attendee’s individual Outlook / Exchange free-busy calendar data.

## Goal
The macro shall:
1. Check each attendee’s free/busy availability individually.
2. Treat the following statuses as BLOCKED:
   - Tentative
   - Busy
   - Out of Office
   - Any other non-free status
3. Find the common free slots for the required attendees.
4. Also check one optional attendee and mark whether that optional attendee is free in the slot.
5. Export the final result to a CSV file saved into:
   C:\temp\
6. The CSV filename must be descriptive (“speaking filename”) and include a timestamp of the query execution.
7. The CSV shall contain only the relevant final candidate slots, sorted with the nearest upcoming slot first.

## Required attendees
Use these attendees as hardcoded defaults in the script:
- Sziva, Laszlo <Laszlo.Sziva@kanadevia-inova.com>
- Maties, Oana <Oana.Maties@kanadevia-inova.com>
- Chavlisek, Georgios <Georgios.Chavlisek@kanadevia-inova.com>
- Zelem, David <David.Zelem@kanadevia-inova.com>
- Balla, Andrej <Andrej.Balla@kanadevia-inova.com>

## Optional attendee
Use this attendee as optional:
- Lorenzoni, Enrico <Enrico.Lorenzoni@kanadevia-inova.com>

## Date range
The query period shall be:
- Start date: today (date when the macro is executed)
- End date: 23.06.2026 inclusive

## Time rules
The macro shall search only within these limits:
- Earliest allowed slot start: 08:00
- Avoid any slot overlapping 12:00–14:00
- Do not allow slots starting at 16:00 or later
- Prefer 60-minute slots
- Also return 30-minute slots as fallback
- Use 30-minute granularity

## Required scheduling logic
- The macro must check each attendee individually via Outlook / Exchange free-busy information.
- Do NOT guess availability from meeting search results.
- Do NOT rely on heuristic assumptions like “mornings are likely free”.
- The decision must be based only on actual free/busy data.
- Treat Tentative as blocked for required attendees.
- Treat Out of Office as blocked.
- Treat all non-free values conservatively as blocked.
- Export only slots where ALL required attendees are free.
- For each exported slot, also evaluate whether the optional attendee is free.
- The earliest valid slot in time must appear at the top of the CSV.
- If both 60-minute and 30-minute slots are found, include both, but clearly mark which are preferred and which are fallback.

## Output file requirements
The macro shall save the result as a CSV file in:
C:\temp\

If the folder does not exist, create it automatically.

The filename must be descriptive and include the timestamp of the query execution.

Example filename:
C:\temp\CommonMeetingSlots_20260609_081500.csv

You may use a better descriptive name if it is even clearer, for example:
C:\temp\CommonMeetingSlots_KVI_20260609_081500.csv

## CSV content requirements
The CSV shall contain ONLY the relevant final candidate slots.
Do NOT export:
- raw free/busy strings
- debug information
- internal intermediate calculations
- per-attendee status matrices unless needed for logic only

Each CSV row shall represent one valid common slot.

The CSV shall contain at least these columns:
- QueryTimestamp
- SlotStart
- SlotEnd
- DurationMinutes
- RequiredAttendeesFree
- OptionalAttendeeFree
- OptionalAttendeeName
- Notes

Rules for values:
- QueryTimestamp = timestamp when macro was executed
- SlotStart = full date and time
- SlotEnd = full date and time
- DurationMinutes = 60 or 30
- RequiredAttendeesFree = always TRUE for exported rows
- OptionalAttendeeFree = TRUE or FALSE depending on Enrico’s availability
- OptionalAttendeeName = Lorenzoni, Enrico
- Notes = one of:
  - Preferred 60 min slot
  - Fallback 30 min slot
  - Optional attendee blocked
  - Preferred 60 min slot; optional attendee free
  - Preferred 60 min slot; optional attendee blocked
  - Fallback 30 min slot; optional attendee free
  - Fallback 30 min slot; optional attendee blocked

## Sorting requirements
- Export rows sorted ascending by SlotStart
- The closest upcoming valid slot must be first
- No grouping is needed if pure chronological sorting already ensures closest slot is on top

## Technical requirements
- Use Outlook VBA
- Use Outlook Recipient.FreeBusy or another Outlook-native per-attendee free/busy approach
- Resolve recipients properly before checking FreeBusy
- Handle unresolved recipients gracefully
- Use robust error handling
- Write clean, modular VBA with functions/subs
- Add explanatory comments
- Make the code directly runnable with minimal edits
- Use a CSV format that is safe to open in Excel
- Use semicolon as delimiter if that is safer for European regional settings; otherwise clearly document delimiter choice in comments
- Format date/time values in a clear sortable format, preferably:
  yyyy-mm-dd HH:nn
- Ensure the CSV writing logic does not corrupt special characters in names

## Free/busy mapping requirements
Handle Outlook free/busy values conservatively.
If using Recipient.FreeBusy, assume:
- 0 = free
- 1 = tentative
- 2 = busy
- 3 = out of office
- any other value = blocked unless clearly documented otherwise

All non-zero values shall be treated as blocked for required attendees.

## Error handling requirements
- If a recipient cannot be resolved, handle this gracefully and document that behavior in comments
- If free/busy cannot be read for a required attendee, do not falsely mark any slot as valid
- If C:\temp cannot be created or written to, show a meaningful message
- Avoid silent failures
- Include clear variable names and structured functions

## Maintainability requirements
Structure the code so that the following can easily be changed later:
- attendee list
- optional attendee
- end date
- earliest start time
- excluded lunch window
- latest allowed start time
- slot durations
- output folder
- filename prefix

## Nice-to-have improvements
If feasible, include these improvements:
- Separate functions for:
  - building attendee list
  - resolving recipients
  - reading free/busy
  - checking whether a slot is allowed by time rules
  - checking whether an attendee is free in a slot
  - finding common slots
  - writing CSV
- A helper function for formatting timestamps in filenames
- A helper function for creating the output folder if missing
- Easy configuration section at top of script

## Required output from you
Provide the following in this order:
1. The full VBA code
2. A short explanation of where in Outlook VBA to paste it
3. A short note about Outlook object library assumptions
4. A short explanation of how the free/busy status mapping is handled
5. A short note about any limitations of Recipient.FreeBusy
6. Make sure the code is directly runnable with minimal edits

## Acceptance criteria
The solution is only acceptable if:
- it checks each attendee individually,
- it exports only the valid common slots,
- the file is saved under C:\temp\ with timestamped descriptive filename,
- the earliest valid slot appears at the top of the CSV,
- both 60-minute and 30-minute slots are handled,
- tentative and OOF are treated as blocked,
- the CSV contains only relevant final candidate slots,
- the code is complete VBA and not pseudocode.
``