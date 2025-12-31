import pandas as pd
import yagmail
import os

YOUR_EMAIL = '' # Your able email address
YOUR_APP_PASSWORD = ''  # App password for your gmail account
SUBJECT = "ABLE Mentor | Tест за самооценка на Вашия ученик"
BODY_TEMPLATE = """
Здравейте, изпращаме ви отговорите от теста за самооценка на Вашия ученик. Надяваме се те да ви дадат още по-пълна информация за характера и нагласата му.
Нямаме търпение да се видим на предстоящите събития.

Светли празници,
екипът на ABLE Mentor 
"""

file_path = 'contacts.csv' # Path to the CSV file with mentor and student information
sheet = pd.read_csv(file_path)
yag = yagmail.SMTP(YOUR_EMAIL, YOUR_APP_PASSWORD)
attachments_folder = 'tests' # Folder containing the documents

for index, row in sheet.iterrows():
    mentor_name = str(row['Ментор име']).strip()
    email = row['Ментор мейл']
    student_name = str(row['Ученик име']).strip()
    matched_file = None
    for file in os.listdir(attachments_folder):
        if student_name in file and "_НЕ" not in file:
            matched_file = os.path.join(attachments_folder, file)
            break
    if not matched_file:
        print(f"[SKIPPED] No matching file found for student: {student_name}")
        continue
    try:
        yag.send(
            to=email,
            subject=SUBJECT,
            contents=BODY_TEMPLATE,
            attachments=matched_file
        )
        print(f"[SENT] {matched_file} to {email} (mentor: {mentor_name})")
    except Exception as e:
        print(f"[ERROR] Failed to send to {email}: {e}")