# txt-to-excel

## Project Overview
This project is designed to convert a text file containing lines of data into an Excel file. Each line in the text file follows the format: "你们可以叫我妹妹吗。 -> 3". The application will separate the text from the corresponding number and write them into two columns in an Excel file.

## File Structure
```
txt-to-excel
├── src
│   ├── main.py
│   └── utils.py
├── requirements.txt
└── README.md
```

## Requirements
To run this project, you will need the following Python libraries:
- pandas
- openpyxl

You can install these dependencies by running:
```
pip install -r requirements.txt
```

## Usage Instructions
1. Place your input text file in the appropriate directory.
2. Modify the `src/main.py` file to specify the path of your input text file.
3. Run the application by executing:
   ```
   python src/main.py
   ```
4. The output Excel file will be generated in the specified output directory.

## Input Format
The input text file should contain lines formatted as follows:
```
你们可以叫我妹妹吗。 -> 3
```
Each line consists of a text segment followed by " -> " and a corresponding number.

## Output
The application will create an Excel file with two columns:
- Column A: Text segment
- Column B: Corresponding number

## License
This project is licensed under the MIT License.