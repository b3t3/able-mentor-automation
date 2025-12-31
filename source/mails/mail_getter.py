import pandas as pd

FILE = "register.xlsx"
MAIN_SHEET = "Season XXIV"   # change this to your main register sheet name
MENTOR_COL = "BU"    # change this to the column letter for mentor names in the main sheet
MENTORS_SHEETS = {
    "Mentors OLD": {"name_col": "E", "email_col": "G"}, # sheet with old mentors
    "Mentors NEW": {"name_col": "F", "email_col": "H"} # sheet with new mentors
}

OUTPUT_FILE = "mentor_emails_ordered.csv"


def get_column_index(col_letters):
    """Convert Excel-style column letters (A, Z, AA, BU...) to 0-based index."""
    col_letters = col_letters.upper()
    index = 0
    for char in col_letters:
        index = index * 26 + (ord(char) - ord('A') + 1)
    return index - 1  # convert to 0-based index

def load_name_email_map(xlsx):
    """Build a dict of mentor name → email from all mentor sheets."""
    mapping = {}

    for sheet, cols in MENTORS_SHEETS.items():
        try:
            df = pd.read_excel(xlsx, sheet_name=sheet)
        except Exception as e:
            print(f"⚠️ Could not read sheet {sheet}: {e}")
            continue

        col_names = df.columns.tolist()

        try:
            name_col = col_names[get_column_index(cols["name_col"])]
            email_col = col_names[get_column_index(cols["email_col"])]
        except IndexError:
            print(f"⚠️ Columns not found in {sheet}, skipping.")
            continue

        for _, row in df[[name_col, email_col]].dropna().iterrows():
            name = str(row[name_col]).strip().lower()
            email = str(row[email_col]).strip()
            if "@" in email:
                mapping[name] = email

    return mapping

def main():
    # Load all mentor name-email pairs
    mentor_map = load_name_email_map(FILE)

    # Load main register sheet
    main_df = pd.read_excel(FILE, sheet_name=MAIN_SHEET)
    cols = main_df.columns.tolist()

    mentor_idx = get_column_index(MENTOR_COL)
    if mentor_idx >= len(cols):
        print(f"❌ Column {MENTOR_COL} not found in main sheet.")
        return

    mentor_colname = cols[mentor_idx]

    mentors = main_df[mentor_colname].dropna().astype(str).tolist()

    ordered_emails = []
    for name in mentors:
        name_clean = name.strip().lower()
        email = mentor_map.get(name_clean, "NO EMAIL")
        ordered_emails.append({"Mentor": name.strip(), "Email": email})

    # Save result
    out_df = pd.DataFrame(ordered_emails)
    out_df.to_csv(OUTPUT_FILE, index=False, encoding="utf-8-sig")

    print(f"✅ Done! Saved {len(out_df)} mentors to '{OUTPUT_FILE}'.")
    missing = [m["Mentor"] for m in ordered_emails if m["Email"] == "NO EMAIL"]
    if missing:
        print(f"⚠️ {len(missing)} mentors missing emails:")
        for m in missing:
            print(" -", m)

if __name__ == "__main__":
    main()