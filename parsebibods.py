import pandas as pd
import ezodf
from pandas_ods_reader import read_ods

import docx
import re
from docx import Document


import argparse
import sys
from datetime import datetime
import os

"""
parse the command line
"""


def argparser(routine):
    # optional input list variables

    example_text_0 = """Usage:

    python3 {} bib_file.bib R""".format(
        routine
    )

    example_text_1 = """Attention:  if NO -n flag is provided, -r and -ne are ignored: usage example:

    python3 {} -p SIR -c US
    python3 {} -p SIR -c Italy  -r Toscana, Piemonte
    python3 {} -p SIR -c Italy  -ne Lombardia, Veneto""".format(
        routine, routine, routine
    )
    usage_text = example_text_0  # + "\n\n" + example_text_1
    parser = argparse.ArgumentParser(
        prog=routine,
        description="Manages bibtex file and formats its content according to different output formats.",
        epilog=usage_text,
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    # mandatory arguments
    parser.add_argument(
        "bib_file",
        action="store",
        # dest='plot_type',
        help="Provide the bib file",
        metavar="BIB_FILE",
        type=str,
        default="",
    )
    parser.add_argument(
        "type",
        action="store",
        # dest='plot_type',
        help="Provide a type. R for ricercatore, T for tecnologo",
        metavar="TYPE",
        choices=["R", "T"],
        type=str,
        default="",
    )

    parser.add_argument(
        "name",
        action="store",
        # dest='plot_type',
        help="Provide the name",
        metavar="NAME",
        type=str,
        default="",
    )

    

    # optional
    parser.add_argument(
        "-v",
        "--verbose",
        action="store_true",
        required=False,
        dest="verbose",
        help='Print verbose information. Default=False"',
        default=False,
    )

    args = parser.parse_args()
    # get the values

    bib_file = args.bib_file
    type = args.type
    name=args.name
    verbose = args.verbose

    return bib_file, type, name, verbose


"""
parse the ods bibfile to create the list of sheets
"""

'''
parse the ods and extract the sheets
'''
def parseods(bib_file, verbose):
    routine = "parseods"
    sheets = []
    doc = ezodf.opendoc(bib_file)
    print(
        "Routine {}. Spreadsheet {} contains {} sheet(s).".format(
            routine, bib_file, len(doc.sheets)
        )
    )
    for sheet in doc.sheets:
        if verbose:
            print("-" * 40)
            print("   Sheet name : '%s'" % sheet.name)
            print("Size of Sheet : (rows=%d, cols=%d)" % (sheet.nrows()-2, sheet.ncols()))
        sheets.append(sheet)
    return sheets


def create_dataframes_from_sheets(sheets, verbose):
    routine = "create_dataframes_from_sheets"
    df_dict = {}
    for sheet in sheets:
        if verbose:
            print(
                "Routine {}. Creating DataFrame from sheet {}.".format(routine, sheet.name)
            )
        try:
            for i, row in enumerate(sheet.rows()):
                # row is a list of cells
                # assume the header is on the first row
                if i == 0:
                    # columns as lists in a dictionary
                    df_dict = {cell.value: [] for cell in row}
                    # create index for the column headers
                    col_index = {j: cell.value for j, cell in enumerate(row)}
                    continue
                for j, cell in enumerate(row):
                    # use header instead of column index
                    df_dict[col_index[j]].append(cell.value)
            
                # and convert to a DataFrame
                #print(df_dict)
            df = pd.DataFrame(df_dict)
        except:
            print("Routine {}. Error in sheet {}, skipping".format(routine, sheet.name))
        #df

'''
create DF from file and sheets
'''
def create_dataframes_from_odsfile(bib_file,sheets, verbose):
    routine = "create_dataframes_from_odsfile"
    dfs_dict={}
    for sheet in sheets:
        if verbose:
            print(
                "Routine {}. Creating DataFrame from sheet {}.".format(routine, sheet.name)
            )
        try:
            df = read_ods(bib_file, sheet.name, headers=True)
            dfs_dict[sheet.name]=df
        except:
            print("Routine {}. Error in sheet {}, skipping".format(routine, sheet.name))
        #df
    return dfs_dict

'''
article in journal
'''
def parse_article(df):
    ret_text=""
    wrong_item=None
    #table structure
    journal_dict={}
    '''
    dft=df.T.loc[['Titolo','Tipo']]
    print(dft[0])
    '''
    
    new_keys_journal=['Tipologia prodotto ','Elenco autori ','Titolo ','Rivista ','Codice identificativo (ISSN) ','anno pubblicazione ',
    'Indice di classificazione ', 'Impact Factor rivista ','ruolo svolto ', 'numero citazioni ','Altre informazioni ']
    journal_table={}
    
    #init dict
    for index, row in df.iterrows():
        #print(int(row['Anno di pubblicazione']))
        dict_key=int(row['Anno di pubblicazione'])
        journal_dict[dict_key]=[]
    #fill the dict
    for index, row in df.iterrows():
        journal_table={}
        #print(int(row['Anno di pubblicazione']))
        
        dict_key=int(row['Anno di pubblicazione'])
        
        #Tipologia prodotto
        ret_text=new_keys_journal[0]+row["Tipo"]
        journal_table[0]=ret_text
        
        #Elenco autori
        ret_text=new_keys_journal[1]+row["Autore/i"]
        journal_table[1]=ret_text
        
        #Titolo
        ret_text=new_keys_journal[2]+row["Titolo"]
        regex = r"\\'"
        test_str = ret_text
        subst = "'"
        ret_text = re.sub(regex, subst, test_str, 0, re.MULTILINE)
        journal_table[2]=ret_text
        
        #Rivista
        rivista_issn=row['Rivista']
        
        # contains issn as well
        regex = r"ISSN:\s[0-9]+-[0-9xX]+"

        issn=re.findall(regex,rivista_issn)[0]
        result=re.split(regex, rivista_issn)
        rivista=result[0]
        # You can manually specify the number of replacements by changing the 4th argument
        #result = re.sub(regex, subst, test_str, 0, re.MULTILINE)
        ret_text=new_keys_journal[3]+rivista
        journal_table[3]=ret_text
        
        ret_text=new_keys_journal[4]+issn
        journal_table[4]=ret_text
        
        #Anno
        ret_text=new_keys_journal[5]+str(int(row["Anno di pubblicazione"]))
        journal_table[5]=ret_text
       
        #Indicizzato da
        if row["Indicizzato da"] is not None:
            ret_text=new_keys_journal[6]+row["Indicizzato da"]
        else:
            ret_text=new_keys_journal[6]
        journal_table[6]=ret_text
        
        #Impact Factor rivista, ruolo svolto, citazioni 7,8,9
        ret_text=new_keys_journal[7]
        journal_table[7]=ret_text
        
        ret_text=new_keys_journal[8]
        journal_table[8]=ret_text
        
        ret_text=new_keys_journal[9]
        journal_table[9]=ret_text
        
        #Altre informazioni. DOI+abstract, URL
        doi=""
        url=""
        abstract=""
        ret_text=new_keys_journal[10]
        if row['DOI'] is not None:
            doi=row['DOI']
        else:
            doi=""
        if row['Abstract'] is not None:
            abstract=row['Abstract']
        else:
            url=""
        if row['URL'] is not None:
            url=row['URL']
        else:
            url=""
        ret_text=ret_text+' '.join([doi,url,abstract])
        journal_table[10]=ret_text
        (journal_dict[dict_key]).append(journal_table)
        

    return journal_dict


'''
contribute w/o isbn
'''
def parse_contribute_no_isbn(df):
    ret_text=""
    wrong_item=None
    #table structure
    noisbn_dict={}

    proceedings_tables={}
    
    new_keys_no_isbn=['Tipologia prodotto ','Titolo ','Descrizione ','Elenco autori ','Ruolo svolto ','anno pubblicazione ','Altre informazioni ']
    
    
    #init dict
    for index, row in df.iterrows():
        #print(int(row['Anno di pubblicazione']))
        dict_key=int(row['Anno di pubblicazione'])
        noisbn_dict[dict_key]=[]
    #fill the dict
    for index, row in df.iterrows():
        proceedings_tables={}
        #print(int(row['Anno di pubblicazione']))
        
        dict_key=int(row['Anno di pubblicazione'])
        
        #Tipologia prodotto
        ret_text=new_keys_no_isbn[0]+row["Tipo"]+ " (senza ISBN)"
        proceedings_tables[0]=ret_text
              
        #Titolo
        ret_text=new_keys_no_isbn[1]+row["Titolo"]
        regex = r"\\'"
        test_str = ret_text
        subst = "'"
        ret_text = re.sub(regex, subst, test_str, 0, re.MULTILINE)
        proceedings_tables[1]=ret_text
 
        #Descrizione
        ret_text=new_keys_no_isbn[2]
        proceedings_tables[2]=ret_text

        #Elenco autori
        ret_text=new_keys_no_isbn[3]+row["Autore/i"]
        proceedings_tables[3]=ret_text
        
        #Ruolo svolto
        ret_text=new_keys_no_isbn[4]
        proceedings_tables[4]=ret_text
        
        #Anno
        ret_text=new_keys_no_isbn[5]+str(int(row["Anno di pubblicazione"]))
        proceedings_tables[5]=ret_text
        
        
        
        #Altre informazioni. DOI+abstract, URL
        doi=""
        url=""
        abstract=""
        ret_text=new_keys_no_isbn[6]
        '''
        if row['DOI'] is not None:
            doi=row['DOI']
        else:
            doi=""
        '''
        if row['Abstract'] is not None:
            abstract=row['Abstract']
        else:
            url=""
        if row['URL'] is not None:
            url=row['URL']
        else:
            url=""
        ret_text=ret_text+' '.join([doi,url,abstract])
        proceedings_tables[10]=ret_text
        (noisbn_dict[dict_key]).append(proceedings_tables)
        

    return noisbn_dict


"""
main
"""
def main():
    new_document = Document()
    routine = sys.argv[0]
    sheets = []
    dfs_dict={}
    dfs_dict_by_year_journal={}
    dfs_dict_by_year_noisbn={}
    df=None
    # output file
    new_doc_name = "./bib/bibfilefromods"
    suffix = ".docx"

    (bib_file, type, name, verbose) = argparser(routine)
    new_doc_name=new_doc_name+"_"+type+"_"+name+suffix
   
    print(bib_file, type, name, verbose)
    sheets = parseods(bib_file, verbose)
    # new_doc_name=new_doc_name+"_"+type+suffix
    # tables_dict,y=parse_bibtext(bib_file, type)
    #print(sheets)
    dfs_dict=create_dataframes_from_odsfile(bib_file,sheets, verbose)
    # print_doc(new_document,new_doc_name,tables_dict)
    for sheet in sheets:
        dfs=[]
        sheet_name=sheet.name
        if sheet_name=='Articolo_in_rivista':
            if verbose:
                print("Calling parse_article with a df with {} rows".format(len(dfs_dict[sheet_name])))
            df=dfs_dict[sheet_name]
            print(df['DOI'])
            dfs_dict_by_year_journal=parse_article(df)
        if sheet_name=='Abstract_in_atti_di_':
            if verbose:
                print("Calling parse_contribute_no_isbn with a df with {} rows".format(len(dfs_dict[sheet_name])))
            df=dfs_dict[sheet_name]
            print(df)
            dfs_dict_by_year_noisbn=parse_contribute_no_isbn(df)
        if sheet_name=='Contributo_in_atti_d':
            if verbose:
                print("Filtering data with ISBN from data with no ISBN from a df with {} rows".format(len(dfs_dict[sheet_name])))
            
            df=dfs_dict[sheet_name]
            df_isbn=df.loc[df['ISBN'] != "\n"]
            df_noisbn=df.loc[df['ISBN'] == "\n"]
            if verbose:
                print("Calling parse_contribute_no_isbn with a df with {} rows".format(len(df_noisbn)))
        
            print(df)
            print(df_isbn)
            print(df_noisbn)
            #dfs_dict_by_year_noisbn=parse_contribute_no_isbn(df)
    #print("XXX ",dfs_dict_by_year_noisbn, len(dfs_dict_by_year_noisbn[2017]))
    
    
    #print("MOISBN=",dfs_dict_by_year_noisbn, "JOURNAL=",dfs_dict_by_year_journal)




main()
