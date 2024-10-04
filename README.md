# Bingo Card Generator

### Motivation

As part of organizing a trivia bingo session for the Society of Asian Scientists and Engineers (SASE) 2024 National Convention, I realized I needed to generate a large number of bingo cards. While there are online bingo card generators available, many of them charge fees, and they don't offer the specific formatting and customization I required for trivia-style bingo. Additionally, creating bingo cards manually is tedious and prone to errors, such as accidentally generating duplicate cards. To avoid these pitfalls and streamline the process, I decided to write a Python script that could generate the cards for me efficiently and with precision.

### What the Code Does

This Python script automatically generates 200 bingo cards in a custom format for a trivia bingo event. Each bingo card is a 5x5 grid of random numbers between 1 and 50, where each column is populated with numbers from a specific range (e.g., column 1 contains numbers from 1-10, column 2 from 11-20, and so on). The grid is formatted with specific dimensions, margins, and fonts, so that the final result fits the requirements of the event. The cells of each bingo card are evenly spaced and bordered to ensure clarity. Once generated, the script creates a Word document (`SASE-Bingo-Cards.docx`) containing all 200 cards, each on a separate page. 

The code also ensures that no two cards are identical, reducing the chance of duplicates, which could be an issue when manually creating bingo cards.

I chose Python for this task due to its versatility and ease of use, especially with libraries like `python-docx` which simplify the process of working with Word documents. Python also allows for fast and error-free automation, perfect for generating large quantities of unique bingo cards.

### How to Run the Code

1. Clone the repository from GitHub:
`git clone https://github.com/rhwang1226/bingo-card-generator.git`


2. Install the necessary dependencies using `pip`:
`pip install python-docx`

3. Run the script:
`python main.py`

The script will generate a Word document named `SASE-Bingo-Cards.docx` in the same directory, containing 200 unique bingo cards formatted and ready to print for the event.
