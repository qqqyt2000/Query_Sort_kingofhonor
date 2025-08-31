def read_txt_file(file_path):
    with open(file_path, 'r', encoding='utf-8') as file:
        lines = file.readlines()
    return lines

def process_lines(lines):
    processed_data = []
    for line in lines:
        if '->' in line:
            text, number = line.split('->')
            processed_data.append((text.strip(), number.strip()))
    return processed_data

def write_to_excel(data, output_file):
    import pandas as pd
    
    df = pd.DataFrame(data, columns=['Text', 'Number'])
    df.to_excel(output_file, index=False)