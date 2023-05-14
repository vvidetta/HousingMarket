import locale
import os
import re
import sys
import pandas


class listing:
    def __init__(self):
        self.file = ""
        self.property_number = -1
        self.date = ""
        self.property_type = ""
        self.bedrooms = -1
        self.contract_type = ""
        self.price = 0.0
        self.location = ""


class parse_state:
    def __init__(self, file: str, lines : list[str], listings: list[listing]):
        self.file = file
        self.lines = lines
        self.listings = listings

        self.line_number = 0
        self.line_count = len(lines)
        self.property_count = -1
        self.property_number = -2
        self.current_listing = None
        self.date = None

        self.try_parse_line = self.try_parse_date_line
        self.move = {
            self.try_parse_date_line: self.try_parse_subject_line,
            self.try_parse_subject_line: self.try_parse_property_line,
            self.try_parse_property_line: self.try_parse_price_line,
            self.try_parse_price_line: self.try_parse_description_line,
            self.try_parse_description_line: self.try_parse_location_line,
            self.try_parse_location_line: self.try_parse_property_line
        }

    def current_line(self) -> str:
        return self.lines[self.line_number]

    def expect_next(self) -> None:
        self.try_parse_line = self.move[self.try_parse_line]

    def try_parse_date_line(self) -> bool:
        date_line = re.compile("^Date: .* (\\d+ \\w\\w\\w \\d\\d\\d\\d)")
        m = date_line.match(self.current_line())
        if not m:
          return False
        self.date = m.group(1)
        return True

    def try_parse_subject_line(self) -> bool:
        subject_line = re.compile(f"^Subject: \w+, (\\d+) new")
        m = subject_line.match(self.current_line())
        if not m:
            return False
        print(f"Email has {m.group(1)} properties")
        self.property_count = int(m.group(1))
        return True

    def try_parse_property_line(self) -> bool:
        property_line = re.compile("^Property (\\d+):")
        m = property_line.match(self.current_line())
        if not m:
            return False
        print(f"File: {self.file}")
        print(f"Property number {m.group(1)}")
        self.property_number = int(m.group(1))
        self.current_listing = listing()
        self.current_listing.file = self.file
        self.current_listing.property_number = self.property_number
        self.current_listing.date = self.date
        return True

    def try_parse_price_line(self) -> bool:
        price_line = re.compile("&pound;(\S+)")
        m = price_line.search(self.current_line())
        if not m:
            return False
        print(f"Price {m.group(1)}")
        self.current_listing.price = locale.atof(m.group(1))
        return True

    def try_parse_description_line(self) -> bool:
        description_line = re.compile("(\\d) bedroom (.*) ((for sale)|(to rent))")
        m = description_line.search(self.current_line())
        if not m:
            return False
        print(f"Bedrooms: {m.group(1)}\nProperty Type: {m.group(2)}\nContract Type: {m.group(3)}")
        self.current_listing.bedrooms = int(m.group(1))
        self.current_listing.property_type = m.group(2)
        self.current_listing.contract_type = m.group(3)
        return True

    def try_parse_location_line(self) -> bool:
        print(f"Location: {self.current_line().strip()}")
        self.current_listing.location = self.current_line()
        self.listings.append(self.current_listing)
        self.current_listing = None
        return True


def parse_email(filename: str, listings: list[listing]) -> None:
    lines = []
    with open(filename) as f:
        lines = f.readlines()

    state = parse_state(filename, lines, listings)
    while state.line_number < state.line_count:
        if state.try_parse_line():
            state.expect_next()
            if state.property_number > state.property_count:
                break
        state.line_number += 1


def parse_emails_in_directory(directory : str) -> list[listing]:
    listings = []
    for file in os.listdir(directory):
        filename = os.fsdecode(file)
        if filename.endswith(".eml"):
            parse_email(f"{directory}\\{filename}", listings)
    return listings


def serialize_listings(output_file: str, listings: list[listing]) -> None:
    records = []
    for x in listings:
        records.append([x.file, x.property_number, x.date, x.price, x.bedrooms, x.property_type, x.contract_type, x.location])
    df = pandas.DataFrame(records, columns=["File", "Property Number", "Date", "Price", "Bedrooms", "Property type", "Contract type", "Location"])
    df.to_excel(output_file)


def main(args : list) -> int:
    if len(args) != 3:
        usage ="""
Usage analyze.py <Directory> <Output>

<Directory> is a path to a directory containing .eml files.
<Output> is a path to the xlsx file to write the results into.

Parses Rightmove Property alert emails into an XLSX file.
"""
        print(usage, file=sys.stderr)
        return 1

    directory = args[1]
    output = args[2]

    locale.setlocale( locale.LC_ALL, 'en_GB.UTF-8' )
    listings = parse_emails_in_directory(directory)
    serialize_listings(output, listings)

    return 0

if __name__ == "__main__":
    sys.exit(main(sys.argv))
