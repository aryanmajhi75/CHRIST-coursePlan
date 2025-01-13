import json
import re

from spire.doc import *
from spire.doc.common import *

# Create a Document object
document = Document()


def extract_text():
    # Load a Word document
    document.LoadFromFile("CoursePlans/Syllabus_MCA272– Programming Using Java.docx")

    # Extract the text of the document
    document_text = document.GetText()

    # Substitutes '\t' with 1 '\n'
    document_text = re.sub(r"\t{1,}", "\n", document_text)
    # Trim the spaces after newline
    document_text = re.sub(r"\n\s+", "\n", document_text)
    # Substitutes '\r' with ''
    document_text = document_text.replace("\r", "")

    with open("Output/DocumentText.txt", "w", encoding="utf-8") as file:
        file.write(document_text)


# JSON Structure:

# [
#   CourseName-Num: ,
#   Total-Teaching-Hours: ,
#   Max-Marks: ,
#   Credits: ,
#   Course-Outcomes: [
#     CO1: ,
#   CO2: ,
#   CO3: ,
#   ],
#   Unit1: [
#     Teaching-hours: ,
#     Title: ,
#     Contents: [],
#     Lab-exercises: [],
#         ],
#   Unit2: [
#     Teaching-hours: ,
#     Title: ,
#     Contents: [],
#     Lab-exercises: [],
#         ],
#   Unit3: [
#     Teaching-hours: ,
#     Title: ,
#     Contents: [],
#     Lab-exercises: [],
#         ],
#   Unit4: [
#     Teaching-hours: ,
#     Title: ,
#     Contents: [],
#     Lab-exercises: [],
#         ],
#   Unit5: [
#     Teaching-hours: ,
#     Title: ,
#     Contents: [],
#     Lab-exercises: [],
#         ],
#   Text-books: [],
#   Web-resources: [],
# ]
# def storeJson():
#     document.LoadFromFile("Output/DocumentText.txt")
#     processed_file = document.GetText()

#     # Initialize storage for extracted data
#     extracted_data = {
#         "CourseName-Num": "",
#         "Total-Teaching-Hours": "",
#         "Max-Marks": "",
#         "Credits": "",
#         "Course-Outcomes": {},
#         "Text-books": [],
#         "Web-resources": []
#     }

#     # Split text into sections by line
#     sections = processed_file.split("\n")

#     unit_data = {}
#     collecting_lab_exercises = False
#     current_unit = None

#     for section in sections:
#         section = section.strip()

#         if not section:
#             continue

#         # Extract course details
#         if section.startswith("MCA"):
#             extracted_data["CourseName-Num"] = section
#         elif section.startswith("Total Teaching Hours for Semester:"):
#             extracted_data["Total-Teaching-Hours"] = section.split(": ")[1]
#         elif section.startswith("Max Marks:"):
#             extracted_data["Max-Marks"] = section.split(": ")[1]
#         elif section.startswith("Credits:"):
#             extracted_data["Credits"] = section.split(": ")[1]

#         # Extract course outcomes
#         elif section.startswith("CO"):
#             co_key, co_value = section.split(": ", 1)
#             extracted_data["Course-Outcomes"][co_key] = co_value

#         # Detect units
#         elif section.startswith("Unit-"):
#             current_unit = section.split(" ")[0].replace("-", "")
#             unit_data[current_unit] = {
#                 "Teaching-hours": 0,
#                 "Title": "",
#                 "Contents": [],
#                 "Lab-exercises": []
#             }
#             collecting_lab_exercises = False

#         # Extract teaching hours
#         elif section.startswith("Teaching Hours:") and current_unit:
#             unit_data[current_unit]["Teaching-hours"] = int(section.split(": ")[1])

#         # Extract unit title
#         elif section.isupper() and current_unit:
#             unit_data[current_unit]["Title"] = section

#         # Collect contents until lab exercises start
#         elif not collecting_lab_exercises and current_unit and not section.startswith("Lab Exercises:"):
#             unit_data[current_unit]["Contents"].append(section)

#         # Start collecting lab exercises
#         elif section.startswith("Lab Exercises:"):
#             collecting_lab_exercises = True

#         # Collect lab exercises
#         elif collecting_lab_exercises and section[0].isdigit():
#             unit_data[current_unit]["Lab-exercises"].append(section)

#         # Extract textbooks and references
#         elif section.startswith("Text Books and Reference Books"):
#             collecting_textbooks = True
#         elif section.startswith("Web Resources"):
#             collecting_textbooks = False
#             collecting_web_resources = True
#         elif collecting_textbooks:
#             extracted_data["Text-books"].append(section)
#         elif collecting_web_resources:
#             extracted_data["Web-resources"].append(section)

#     # Merge unit data into extracted data
#     extracted_data.update(unit_data)

#     print(extracted_data)

#     # Write the extracted data into a JSON file
#     with open("Output/DocumentText.json", "w", encoding="utf-8") as file:
#         json.dump(extracted_data, file, ensure_ascii=False, indent=4)


def storeJson():
    with open("Output/DocumentText.txt", "r", encoding="utf-8") as file:
        processed_file = file.read()

    # Initialize storage for extracted data
    extracted_data = {
        "CourseName-Num": "",
        "Total-Teaching-Hours": "",
        "Max-Marks": "",
        "Credits": "",
        "Course-Outcomes": {},
        "Units": {},
        "Text-books": [],
        "Web-resources": [],
    }

    # Split text into sections by line
    sections = processed_file.split("\n")

    unit_data = {}
    collecting_lab_exercises = False
    collecting_textbooks = False
    collecting_web_resources = False
    current_unit = None

    for section in sections:
        section = section.strip()

        if not section:
            continue

        # Extract course details
        if "MCA" in section:
            extracted_data["CourseName-Num"] = section
        elif section.startswith("Total Teaching Hours for Semester:"):
            extracted_data["Total-Teaching-Hours"] = section.split(": ")[1]
        elif section.startswith("Max Marks:"):
            extracted_data["Max-Marks"] = section.split(": ")[1]
        elif section.startswith("Credits:"):
            extracted_data["Credits"] = section.split(": ")[1]

        # Extract course outcomes
        elif section.startswith("CO") and ":" in section:
            co_key, co_value = section.split(": ", 1)
            extracted_data["Course-Outcomes"][co_key] = co_value

        # Detect units
        elif section.startswith("Unit-"):
            current_unit = section.split("-")[1]
            unit_data[current_unit] = {
                "Teaching-hours": 0,
                "Title": "",
                "Contents": [],
                "Lab-exercises": [],
            }
            collecting_lab_exercises = False

        # Extract teaching hours
        elif section.startswith("Teaching Hours:") and current_unit:
            unit_data[current_unit]["Teaching-hours"] = int(section.split(": ")[1])

        # Extract unit title (Upper case assumption)
        elif section.isupper() and current_unit and not collecting_lab_exercises:
            unit_data[current_unit]["Title"] = section

        # Collect contents until lab exercises
        elif current_unit and not collecting_lab_exercises and section:
            unit_data[current_unit]["Contents"].append(section)

        # Start collecting lab exercises
        elif section.startswith("Lab Exercises:"):
            collecting_lab_exercises = True
            unit_data[current_unit]["Lab-exercises"].append(section)

        # Extract textbooks and references
        elif "Text Books and Reference Books" in section:
            # collecting_textbooks = True
            # collecting_web_resources = False
            extracted_data["Text-books"].append(section)

        elif "Web Resources:" in section:
            # collecting_textbooks = False
            # collecting_web_resources = True
            extracted_data["Web-resources"].append(section)

    # Merge unit data into extracted data
    extracted_data.update(unit_data)

    # Write extracted data to JSON file
    with open("Output/DocumentText.json", "w", encoding="utf-8") as file:
        json.dump(extracted_data, file, ensure_ascii=False, indent=4)

    print(json.dumps(extracted_data, indent=4))


# Write the extracted text into a text file
extract_text()
storeJson()

document.Close()
