/*
 * What a REAL filled Iowa DOT Travel Authorization form extracts to.
 *
 * Verbatim output of attachtext.pdfText() on the PDF Matt filled in, kept as
 * a fixture because two things about it are not guessable and both cost a
 * whole afternoon:
 *
 *   - every field object lives inside a compressed /ObjStm, so a scan of the
 *     raw bytes finds "/Widget" zero times, and
 *   - the field NAMES lost their ligatures when whoever built the form let
 *     Acrobat auto-name the boxes from the printed page. "Location" is
 *     spelled "Locaton" in the file itself.
 */
"use strict";
module.exports = "Name & Cost Center: Matthew Miller 30000\nOther Staf Atending: Cedric Wilkinson, Brian Worrel, Tony Gustafson, Chris Pelton, Deanne Popp\nName of Conference: EDC-8 Midwest Peer Exchange\nLocaton: Hyatt Regency 315 Chestnut St, St. Louis, MO 63102\nConference Dates & Times: 10/15/26 7:00 am to 5:00 pm\nDeparture Date: 10/14/2026\nReturn Date: 10/15/2026\nText9: Matt is lead for Connected Corridors for Iowa on Every Day Counts round 8. Matt is pooled fund study champion for the Connected Corridors Advancement Initiative and is directly at the center of this portion of EDC-8.\nCost of Travel Mode miles, estmated fight cost\\: 770 miles\nLuggage Fees: $: 0\nParking: $: 40\nLodging: 1\nnights @ $: 176.90\n= $: 176.90\nRegistraton Fee: $: 0\nTaxi/Uber Fees: $: 0\nName: FHWA\nMaximum Reimbursement Amount: $: N/A\nCheck Box2: checked\nCheck Box5: checked";
