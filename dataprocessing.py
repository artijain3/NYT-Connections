import pandas as pd
import os
import logging
import json

def excel_to_txt(excel_file_path, output_file_name, output_file_dir):
    """
    This function will take in an input data spreadsheet and output a .txt file with the data. 
    This is to organize the data before tokenization.
    The output will be split by category: synonyms, pop_culture, thematic, word_based, fitb (fill in the blank)
    
    Inputs:
        excel_file_path - this is the path to the original excel spreadsheet with the data
        output_file_name - this is the name of the output text file
        output_file_dir - this is the name of the directory where the output text file will be saved
        
    Returns:
        None
        
    Output File Format:
        SYNONYMS = {
            "These are my four associated words: badgers, bugs, hounds, nags": "pester",
    """
    if not os.path.exists(excel_file_path):
        raise FileNotFoundError(f"The file, {excel_file_path} does not exist.")
    
    if not os.path.exists(output_file_dir):
        logging.debug(f"Output file dir: {output_file_dir} does not exist. Creating file in {os.getcwd()}/preprocessed_data")
        os.mkdir(f"{os.getcwd()}/preprocessed_data")
        output_file = os.path.join(f"{os.getcwd()}/preprocessed_data", f"{output_file_name}.txt")
    else:
        output_file = os.path.join(output_file_dir, f"{output_file_name}.txt")
        
    df = pd.read_excel(excel_file_path)
    
    categories = ["SYNONYMS", "POP_CULTURE", "THEMATIC", "WORD_BASED", "FITB"]
    category_dictionary = {}
    for _, row in df.iterrows():
        if row[0] in categories:        
            if category_dictionary != {}:
                with open(output_file, 'a') as outfile:
                    outfile.write(f"{heading} = {json.dumps(category_dictionary, indent=2)}\n")
            category_dictionary = {}
            heading = row[0]
        else:
            string_split = row[0].split('"')
            string_split = string_split[1].split(":")
            key = row[1].split('"')
            key = key[1].lower()
            category_dictionary[f"These are my four associated words:"+string_split[1].lower()] = key
            
    # last category
    with open(output_file, 'a') as outfile:
        outfile.write(f"{heading} = {json.dumps(category_dictionary, indent=2)}\n")
            
    logging.info(f"Output file written to: {output_file}")

            
        
if __name__ == "__main__":
    excel_to_txt("./original_dataset.xlsx", "originaldata", "./preprocessed_data")
        