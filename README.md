# Excel Text to Speech: English Sentence Construction Practice for Persian Speakers

## Overview

"Excel Text to Speech" is a Python application designed to help Persian speakers practice building English sentences. The application presents a Persian sentence, giving the user a chance to construct the English equivalent. After a short delay, the application speaks the correct English sentence aloud and displays it, allowing the user to compare their attempt with the accurate translation and reinforce their learning. This method supports vocabulary acquisition, grammar practice, and improved sentence structure.

## Key Features

*   **Persian-to-English Learning:** Specifically designed for Persian speakers learning English.
*   **Excel Integration:** Reads both Persian sentences and their English translations from `.xlsx` or `.xls` files.
*   **Dual Text Box Display:**
    *   Displays the Persian sentence in one text box, prompting the user to create the English translation.
    *   After a delay, reveals the correct English sentence in a separate text box for comparison.
*   **Text-to-Speech:** Speaks the English sentence aloud using the `pyttsx3` library, reinforcing pronunciation.
*   **Customizable Speech Speed:** Adjust the speaking rate to suit the user's listening comprehension.
*   **Go-To-Cell Functionality:** Jump to a specific sentence pair in your Excel sheet for focused practice.
*   **Randomized Playback:** Present sentences in a random order to challenge recall and prevent rote memorization.
*   **Adjustable Initial Delay:** Customize the delay between the Persian sentence display and the revealing of the English translation, allowing sufficient time for mental translation.
*   **Pause/Resume:** Easily pause and resume the learning session.
*   **User-Friendly GUI:** Built with `tkinter` for a simple, intuitive, and distraction-free learning environment.

## Screenshots (Optional)

*You can add screenshots of the application here to showcase its interface and demonstrate the Persian and English text boxes.*

## Installation

1.  **Prerequisites:** Ensure Python 3.6 or higher is installed on your system.

2.  **Clone the Repository:**

    ```
    git clone [YOUR_REPOSITORY_URL]  # Replace with your repository URL
    cd ExcelTextToSpeech
    ```

3.  **Install Dependencies:**

    ```
    pip install -r requirements.txt
    ```

    *Alternatively, install the libraries individually:*

    ```
    pip install pandas pyttsx3 tkinter
    ```

## Usage

1.  **Prepare Your Excel File:**

    *   Create an Excel file (e.g., `persian_english_sentences.xlsx`) with two columns: one for the Persian sentence and one for its English translation.
    *   The first row should contain column headers.

    ```
    | Persian Sentence        | English Translation        |
    |-------------------------|-----------------------------|
    | سلام، حالت چطوره؟       | Hello, how are you?         |
    | امروز هوا خیلی خوبه.  | The weather is great today. |
    | ...                     | ...                         |
    ```

2.  **Run the Application:**

    ```
    python updatesayexcel.py
    ```

3.  **Using the GUI:**

    *   **Select Excel File:** Click "Browse" to open your Excel file.
    *   **Select Sheet:** Choose the sheet containing your sentence pairs.
    *   **Select Columns:** Designate the column containing the Persian sentences as the "Translation Text" and the English sentences as the "Current Text".
    *   **Set Initial Delay:** Adjust the delay (in seconds) before the English translation is displayed and spoken. This delay is the time you have to construct the English sentence in your mind.
    *   **Go To Cell (Optional):** Enter a cell reference (e.g., "A2") to start from a specific sentence pair.
    *   **Random Order (Optional):** Enable "Random Order" to shuffle the sentence presentation for more challenging practice.
    *   **Adjust Speed:** Control the speaking speed using the slider.
    *   **Start:** Click "Start" to begin the learning session.
    *   **Pause/Resume:** Use "Pause" and "Resume" to control the session flow.
    *   **End:** Click "End" to terminate the session.

## Example `requirements.txt`

