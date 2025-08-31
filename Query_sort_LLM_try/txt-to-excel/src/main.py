import pandas as pd

def read_txt_file(file_path):
    with open(file_path, 'r', encoding='utf-8') as file:
        lines = file.readlines()
    return lines

def process_lines(lines):
    data = []
    for line in lines:
        if '->' in line:
            text, number = line.split('->')
            text = text.strip()
            number = number.strip()
            data.append((text, number))
    return data

def write_to_excel(data, output_file):
    df = pd.DataFrame(data, columns=['Text', 'Number'])
    df.to_excel(output_file, index=False)

def main():
    input_file = 'D:\query_part.txt'  # Update with your input file path
    output_file = 'output.xlsx'  # Desired output Excel file name

    lines = read_txt_file(input_file)
    processed_data = process_lines(lines)
    write_to_excel(processed_data, output_file)

if __name__ == '__main__':
    main()