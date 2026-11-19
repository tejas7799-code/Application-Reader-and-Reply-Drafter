Application Reader and Reply Drafter

Reads all 50 applications from applications.csv, scores each for completeness (0-10) with a reason, and drafts a personal reply email so a human only has to check and send.

Run
python reader.py applications.csv

Python 3 standard library only. openpyxl is optional (used for output.xlsx). No API key, no network.

Outputs:

output.csv / output.xlsx: score, reason, human-check flag, drafted subject and reply for every application.
review.html: open in any browser. Filter, edit each draft, copy it, or download the edited CSV.
What I built
Clean the data: trims and title-cases names, maps college spellings to one name (MBU, "Mohan Babu Univeristy" -> Mohan Babu University), parses the two date formats, treats N/A, Not provided and blanks as missing, normalises +91 phone numbers.
Score (start at 10): resume missing -3, statement missing -4, statement too short or generic ("interested", "to gain knowledge") -3, invalid email -3, invalid phone -2. The reason column lists every deduction.
Draft a reply from a template using name, college and domain. It asks only for what is missing (resume, a specific statement, a correct phone number).

Bonus items done:

Flag for human check: duplicate, invalid email (cannot be emailed, contact by phone), score of 4 or below, or statement that points to a different domain than the one chosen (the reply then asks them to confirm).
Duplicates: same email = same person. The latest submission is kept; earlier ones are marked "do not send" with no reply. Found: Hema Sharma (1004/1048), Sahithi Rao (1049/1022), Rohit Kumar (1047/1050).
Review screen: review.html.
Cost: 0 API calls, 0 cost, because everything is rules and templates.
Choices
Rules, not an AI model. With 50 rows and a fixed checklist, rules are free, instant, repeatable and every score can be explained. The drawback is that replies are templated and the statement check is keyword based.
Latest duplicate wins, since people usually resubmit to fix something.
"SVEC" is assumed to mean Sree Vidyanikethan Engineering College.
A mismatch between statement and domain is flagged, not penalised, as it is a possible data issue rather than missing information.
With more time
Use a small local or free-API model only for the personal sentence about the statement, with rules still doing scoring.
Check that resume links actually open, and detect near-duplicates (same phone, different email).
Unit tests for the cleaning functions and a send-via-SMTP step after review.
Where I used AI tools

I used Claude to plan the scoring rules and write the first version of the script and this README, then ran it on the dataset and checked the output.
