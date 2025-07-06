import pandas as pd

OXFORD_DB_FILE = "oxford_db.xlsx"

def import_oxford_db(filename="oxford_db.xlsx"):
    """
    Import the Oxford database dictionary.
    Returns:
        dict: A dictionary with words as keys and their definitions as values.
    """
    df_oxford_db = pd.read_excel(OXFORD_DB_FILE, engine="openpyxl", header=None)

    # create a dictionary which key is the first column and value is the second column
    oxford_db_dict = {}
    for index, row in df_oxford_db.iterrows():
        first_col_value = row.iloc[0]
        second_col_value = row.iloc[1]
        oxford_db_dict[first_col_value] = second_col_value

    return oxford_db_dict

def save_word_to_oxford_db(word, definition, filename=OXFORD_DB_FILE):
    """
    Save a word and its definition to the Oxford database.
    
    Args:
        word (str): The word to save.
        definition (str): The definition of the word.
        filename (str): The name of the file to save the database to.
    """
    df_oxford_db = pd.read_excel(filename, engine="openpyxl", header=None)
    new_row = pd.DataFrame([[word, definition]])
    df_oxford_db = pd.concat([df_oxford_db, new_row], ignore_index=True)
    df_oxford_db.to_excel(filename, index=False, header=False)