# Weekly Statistics Email Sender

This C# program extracts data from an Excel sheet using Aspose and initializes an SMTP server for Outlook to send weekly statistics emails to all specified users.

![example picture](https://github.com/mariekart42/Wochenstatistik/blob/main/example.png)

## Features

- Extracts data from an Excel file (`Daten Wochenstatistik.xlsx`).
- Reads a list of users from a text file (`Nutzer Liste.txt`).
- Sends personalized emails with the extracted statistics to each user.

## Prerequisites

- .NET SDK installed on your system.
- Aspose.Cells library for handling Excel files.
- Access to an SMTP server for sending emails.

## Setup Instructions

1. ### Clone the Repository
   ```bash
   git clone https://github.com/mariekart42/Wochenstatistik.git

---
2. ### Insert Required Files
- In the document folder, insert the following files:
    - `Daten Wochenstatistik.xlsx`: This file should be an Excel sheet containing the up-to-date values required for a correct weekly statistic.
    - `Nutzer Liste.txt`: This file contains all users who expect to receive an email with their personal statistics.
---
3. ### File Format Specifications
- **Daten Wochenstatistik.xlsx**:
    - Ensure this file is a valid Excel sheet with the required data for the weekly statistics.
- **Nutzer Liste.txt**:
    - This file should contain the email addresses and user initials of the recipients in the following format:
      `example@eisenfuhr.com|MSG`
    - Note:
        - Each line must contain a valid email address and user initials separated by a |.
        - Empty lines or lines starting with # will be ignored.
        - The program will terminate if the syntax is incorrect and no emails will be sent.
---
4. ### Build and Run the Program
- Restore dependencies:
  `dotnet restore`
- Build the project:
  `dotnet build -c Release`
- Publish the project:
  `dotnet publish -c Release -o ./publish`
- Run the make command:
  `make`
- Run the program:
  `dotnet ./publish/WeeklyStatisticsEmailSender.dll`
---
5. ### Notes
- The program will not work if file names are misspelled, the file type is wrong, or if the file is not provided.
- Ensure that the email addresses in the Nutzer Liste.txt are correct and reachable via the SMTP server configured in the program.
