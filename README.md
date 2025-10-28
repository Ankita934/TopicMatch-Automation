📘 TopicMatch – Automated Syllabus Checker
🧠 Overview

TopicMatch is a Python-based automation project that checks whether every topic listed in a syllabus is covered in your teacher’s Word notes (.docx files).
It uses fuzzy matching to detect even slight word differences and generates an Excel report showing which topics are found or missing — saving hours of manual checking!

⚙️ Features

🔍 Scans multiple .docx Word files at once

🧾 Matches topics with partial and fuzzy similarity

📊 Exports an Excel report with match scores and snippets

⚡ Handles punctuation, case, and spacing differences

💼 Built with Python libraries: pandas, rapidfuzz, and python-docx

🧩 Tech Stack

Language: Python

Libraries: Pandas, RapidFuzz, python-docx, openpyxl

🧠 How It Works

Reads all syllabus topics from a syllabus.txt file (one topic per line).

Reads all .docx Word notes from the specified folder.

Compares each topic across all documents using fuzzy matching.

Generates an Excel file (syllabus_check_report.xlsx) showing:

Topic name

Whether it was found or not

Match score (0–100%)

Which Word file matched

A small snippet of matched text
