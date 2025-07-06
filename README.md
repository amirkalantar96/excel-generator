# 📊 Excel Generator

## Description of the project
- Instantly create a pre-formatted Excel file with a single terminal command.
- Saves time by skipping manual Excel file creation and column formatting.
- Initially built for personal use to streamline repetitive Excel tasks.
- The settings (such as columns, structure, etc.) are predefined in code and customizable
- Especially useful for developers who need fast, consistent Excel files.
- Open for future improvements and general user interfaces.

## How does it work?
1. Open Terminal.
2. Change the current working directory to the location where you want the cloned directory.
3. Copy bellow command in to terminal
  ```
  $ git clone https://github.com/amirkalantar96/excel-generator.git
  ```
4. Press Enter to create your local clone.
  ```
  Cloning into 'excel-generator'...
  ```
5. Install the project dependencies by running the following command inside the cloned directory:
  ```
  npm install
  ```
6. You can customize the settings and names according to your needs inside the index.js file.
7. Run the following command in the terminal to start the project:
  ```
  npm run buildExcel
  ```
8. You will be prompted with the following question in the terminal:
  ```
  Press Enter to create the Excel file:
  ```
Type your desired filename and press Enter.<br/>
The file will be generated accordingly.

> [!NOTE]
> This project currently works via CLI (Command Line Interface).<br/>
> No graphical user interface is implemented yet.

## 🧰 Tech Stack

- **Runtime Environment:**  
  ![Node.js](https://img.shields.io/badge/Node.js-339933?style=flat&logo=node.js&logoColor=white)

- **Library:**  
  ![ExcelJS](https://img.shields.io/badge/ExcelJS-4.4.0-blue?style=flat&logo=Microsoft%20Excel&logoColor=white)