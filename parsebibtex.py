#!/usr/bin/python3

import bibtexparser
from bibtexparser.bwriter import BibTexWriter
from bibtexparser.bibdatabase import BibDatabase

import docx
import re
from docx import Document


import argparse
import sys
from datetime import datetime
import os


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

    args = parser.parse_args()
    # get the values

    bib_file = args.bib_file
    type = args.type

    return bib_file, type

def parse_article(n, bib_item):
    ret_text=""
    #table structure
    new_keys_journal=['Nr ', 'Tipologia prodotto ','Elenco autori ','Titolo ','Rivista ','Codice identificativo (ISSN) ','anno pubblicazione ',
    'Indice di classificazione ', 'Impact Factor rivista ','ruolo svolto ', 'numero citazioni ','Altre informazioni ']
    journal_table={}
    #id
    ret_text=new_keys_journal[0]+str(n)
    journal_table[0]=ret_text
    
    #Tipologia prodotto
    ret_text=new_keys_journal[1]+"Articolo in Rivista"
    journal_table[1]=ret_text
    
    #Elenco autori
    ret_text=new_keys_journal[2]+bib_item['author']
    journal_table[2]=ret_text
    
    #Titolo
    ret_text=new_keys_journal[3]+bib_item['title']
    regex = r"\\'"
    test_str = ret_text
    subst = "'"
    ret_text = re.sub(regex, subst, test_str, 0, re.MULTILINE)
    journal_table[3]=ret_text
    
    #Rivista
    ret_text=new_keys_journal[4]+bib_item['journal']
    journal_table[4]=ret_text
    
    #ISSN
    ret_text=new_keys_journal[5]+bib_item['issn']
    journal_table[5]=ret_text
    
    #anno
    ret_text=new_keys_journal[6]+bib_item['year']
    journal_table[6]=ret_text
    
    #fields from 7 to 10
    ret_text=new_keys_journal[7]
    journal_table[7]=ret_text
    
    ret_text=new_keys_journal[8]
    journal_table[8]=ret_text
    
    ret_text=new_keys_journal[9]
    journal_table[9]=ret_text
    
    ret_text=new_keys_journal[10]
    journal_table[10]=ret_text
    
    #managing altre informazioni
    # I put url, doi, abstract
    ret_text=new_keys_journal[11]+"\n"
    doi=""
    url=""
    abs=""
    try:
        doi="doi: "+bib_item['doi']
    except:
        print("item, ",n, "no doi")

    try:
        url="url: "+bib_item['url']
    except:
            print("item, ", n, "no url")
    try:
        abs="abstract: "+bib_item['abstract']
    except:
        print("item, ",n, "no abstract")
    ret_text=ret_text+' '.join([doi,url,abs])
    journal_table[11]=ret_text
    return journal_table

def parse_inproceedings(bib_item):
    return

def parse_bibtext(bib_file):
    tables_dict={}
    temp={}
    with open(bib_file) as bibtex_file:
        bib_database = bibtexparser.load(bibtex_file)

    n = 1
    for bib_item in bib_database.entries:
        print()
        if bib_item['ENTRYTYPE']=="inproceedings":
            print("call inproceedings ",n)
        if bib_item['ENTRYTYPE']=="article":
            print("call article ",n)
            temp=parse_article(n,bib_item)
            tables_dict[n]=temp
        if bib_item['ENTRYTYPE']=="techreport":
            print("call techreport ",n)
        n = n + 1
    print("#items ",n)
    return tables_dict

def main():
    routine = sys.argv[0]
    #output file
    new_doc_name='../bib/test.docx'
    (bib_file, type) = argparser(routine)
    print(bib_file, type)
    x=parse_bibtext(bib_file)
    print(x)

main()
