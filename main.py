# This is a sample Python script.

# !/usr/bin/env python3
"""
Module Docstring
"""

__author__ = "MiKe Howard"
__version__ = "0.1.0"
__license__ = "MIT"


import logging
from logzero import logger
import pandas as pd
import glob
import os
import datetime as dt
import numpy as np


# OS Functions
def filesearch(word=""):
    """Returns a list with all files with the word/extension in it"""
    logger.info('Starting filesearch')
    file = []
    for f in glob.glob("*"):
        if word[0] == ".":
            if f.endswith(word):
                file.append(f)

        elif word in f:
            file.append(f)
            # return file
    logger.debug(file)
    return file

def Change_Working_Path(path):
    # Check if New path exists
    if os.path.exists(path):
        # Change the current working Directory
        try:
            os.chdir(path)  # Change the working directory
        except OSError:
            logger.error("Can't change the Current Working Directory", exc_info = True)
    else:
        print("Can't change the Current Working Directory because this path doesn't exits")

#Pandas Functions
def Excel_to_Pandas(filename,check_update=False,SheetName=None):
    logger.info('importing file ' + filename)
    df=[]
    if check_update == True:
        timestamp = dt.datetime.fromtimestamp(Path(filename).stat().st_mtime)
        if dt.datetime.today().date() != timestamp.date():
            root = tk.Tk()
            root.withdraw()
            filename = filedialog.askopenfilename(title =' '.join(['Select file for', filename]))

    try:
        df = pd.read_excel(filename, SheetName)
        df = pd.concat(df, axis=0, ignore_index=True)
    except:
        logger.error("Error importing file " + filename, exc_info=True)

    df=Cleanup_Dataframe(df)
    logger.debug(df.info(verbose=True))
    return df

def Cleanup_Dataframe(df):
    logger.info('Starting Cleanup_Dataframe')
    logger.debug(df.info(verbose=True))
    # Remove whitespace on both ends of column headers
    df.columns = df.columns.str.strip()

    # Replace whitespace in column header with _
    df.columns = df.columns.str.replace(' ', '_')

    return df

#def TPMO_Work_Categorization():
    #

def main():

    """ Main entry point of the app """
    logger.info("District Resource Plan Main Loop")
    Change_Working_Path('./Data')
    all_files = glob.glob('*.xlsx')
    ResourceDF = pd.DataFrame()

    APDRfilename = 'All Project Data Report _ Current Year.xlsx'
    AllPatfilename = 'All PAT Project Data Report - BIG Kahuna.xlsx'

    try:
        APDR_DF = ResourceDF.append(Excel_to_Pandas(APDRfilename, False))
    except:
        logger.error('Can not find Project Data file')
        raise

    try:
        AllPat_DF = ResourceDF.append(Excel_to_Pandas(AllPatfilename, False))
    except:
        logger.error('Can not find Project Data file')
        raise


    for filename in all_files:
        if filename != APDRfilename and filename != AllPatfilename:
            try:
                ResourceDF = ResourceDF.append(Excel_to_Pandas(filename, False))
            except:
                logger.error('Can not find Project Data file')
                raise

    AllPat_DF['WA_Contractor_Estimate'] = AllPat_DF['Approved_WA_Primoris_Labor'] + AllPat_DF['Approved_WA__Great_Southwestern'] + AllPat_DF['Approved_WA_Other_Cotract_Construction_Labor']

    ResourceDF = pd.merge(ResourceDF,APDR_DF[['PETE_ID', 'REFERENCENUMBER', 'PROJECTCATEGORY', 'REGIONNAME', 'District',
                                              'PGMMGRLASTNAME', 'PROJECTTYPE']], on='PETE_ID', how='left')
    ResourceDF = pd.merge(ResourceDF,AllPat_DF[['PETE_ID', 'WA_Contractor_Estimate']], on='PETE_ID', how='left')



    ResourceDF['Conduit'] = np.nan
    ResourceDF['Grounding'] = np.nan
    ResourceDF['Cable Pulling'] = np.nan
    ResourceDF['Foundations'] = np.nan
    ResourceDF['Below Grade Demo'] = np.nan

    ResourceDF=ResourceDF.replace({'District will do all': 'District', 'District will do some': 'District and Contractor' })
    ResourceDF=ResourceDF.rename(columns={'Comments':'District Comments'})

    #for name in [col for col in ResourceDF.columns if
     #            ResourceDF[col].astype(str).str.contains("Outside District").any()]:
      #  ResourceDF[name] = np.where(ResourceDF['District'] == 'NORTH DALLAS', 'Contractor', ResourceDF[name])

    for name in list(ResourceDF.columns):
        ResourceDF[name] = np.where((ResourceDF['District'] == 'NORTH DALLAS') &
                                    (ResourceDF[name] == 'Outside District')
                                    , 'Contractor', ResourceDF[name])

    ResourceDF['TPMO Work Categorization'] = np.nan
    ResourceDF['TPMO Work Categorization'] = np.where((ResourceDF['Build_Lattice'] == 'District and Contractor') |
                                                      (ResourceDF['FCC'] == 'District and Contractor') |
                                                      (ResourceDF['Install_Insulators'] == 'District and Contractor') |
                                                      (ResourceDF['Build_Lattice'] == 'District and Contractor') |
                                                      (ResourceDF['Replace_Arms'] == 'District and Contractor') |
                                                      (ResourceDF['Set_Switches'] == 'District and Contractor') |
                                                      (ResourceDF['Replace_Arms'] == 'District and Contractor') |
                                                      (ResourceDF['Above_Grade_Demo'] == 'District and Contractor') |
                                                      (ResourceDF['Install_Jumpers'] == 'District and Contractor') |
                                                      (ResourceDF['Remove_Old_Breakers'] == 'District and Contractor') |
                                                      (ResourceDF['Set_Breakers'] == 'District and Contractor') |
                                                      (ResourceDF['Set_Steel'] == 'District and Contractor') |
                                                      (ResourceDF['Weld_Bus'] == 'District and Contractor')
                                                      , 'District and Contractor',
                                                      ResourceDF['TPMO Work Categorization'])




    ResourceDF['TPMO Work Categorization'] = np.where((ResourceDF['Build_Lattice'].isin(['District', np.nan, 'Outside District'])) &
                                                      (ResourceDF['FCC'].isin(['District', np.nan, 'Outside District'])) &
                                                      (ResourceDF['Install_Insulators'].isin(['District', np.nan, 'Outside District'])) &
                                                      (ResourceDF['Build_Lattice'].isin(['District', np.nan, 'Outside District'])) &
                                                      (ResourceDF['Replace_Arms'].isin(['District', np.nan, 'Outside District'])) &
                                                      (ResourceDF['Set_Switches'].isin(['District', np.nan, 'Outside District'])) &
                                                      (ResourceDF['Replace_Arms'].isin(['District', np.nan, 'Outside District'])) &
                                                      (ResourceDF['Above_Grade_Demo'].isin(['District', np.nan, 'Outside District'])) &
                                                      (ResourceDF['Install_Jumpers'].isin(['District', np.nan, 'Outside District'])) &
                                                      (ResourceDF['Remove_Old_Breakers'].isin(['District', np.nan, 'Outside District'])) &
                                                      (ResourceDF['Set_Breakers'].isin(['District', np.nan, 'Outside District'])) &
                                                      (ResourceDF['Set_Steel'].isin(['District', np.nan, 'Outside District'])) &
                                                      (ResourceDF['Weld_Bus'].isin(['District', np.nan, 'Outside District']))
                                                      , 'District',
                                                      ResourceDF['TPMO Work Categorization'])

    ResourceDF['TPMO Work Categorization'] = np.where((ResourceDF['Build_Lattice'].isin(['Contractor', np.nan])) &
                                                      (ResourceDF['FCC'].isin(['Contractor', np.nan])) &
                                                      (ResourceDF['Install_Insulators'].isin(['Contractor', np.nan])) &
                                                      (ResourceDF['Build_Lattice'].isin(['Contractor', np.nan])) &
                                                      (ResourceDF['Replace_Arms'].isin(['Contractor', np.nan])) &
                                                      (ResourceDF['Set_Switches'].isin(['Contractor', np.nan])) &
                                                      (ResourceDF['Replace_Arms'].isin(['Contractor', np.nan])) &
                                                      (ResourceDF['Above_Grade_Demo'].isin(['Contractor', np.nan])) &
                                                      (ResourceDF['Install_Jumpers'].isin(['Contractor', np.nan])) &
                                                      (ResourceDF['Remove_Old_Breakers'].isin(['Contractor', np.nan])) &
                                                      (ResourceDF['Set_Breakers'].isin(['Contractor', np.nan])) &
                                                      (ResourceDF['Set_Steel'].isin(['Contractor', np.nan])) &
                                                      (ResourceDF['Weld_Bus'].isin(['Contractor', np.nan]))
                                                      , 'Contractor',
                                                      ResourceDF['TPMO Work Categorization'])

    items=['Build_Lattice', 'Install_Insulators', 'Build_Lattice', 'Replace_Arms',  'Above_Grade_Demo', 'Install_Jumpers', 'Remove_Old_Breakers', 'Set_Breakers','Set_Steel', 'Weld_Bus' ]

    for item in items:
        for itemm in items:
            ResourceDF['TPMO Work Categorization'] = np.where((ResourceDF['TPMO Work Categorization'] == 'nan') &
                                                       (~ResourceDF[item].isin([ResourceDF[itemm], np.nan]))
                                                       ,  'District and Contractor',
                                                      ResourceDF['TPMO Work Categorization'])



    ResourceDF = ResourceDF[['PETE_ID',
                             'WA',
                             'REGIONNAME',
                             'District',
                             'PGMMGRLASTNAME',
                             'REFERENCENUMBER',
                             'PROJECTCATEGORY',
                             'WA_Contractor_Estimate',
                             'TPMO Work Categorization',
                             'Project_Name',
                             'Description',
                             'Construction_Ready',
                             'Estimated_In-Service_Date',
                             'Budget_Item',
                             'Build_Lattice',
                             'FCC',
                             'Install_Insulators',
                             'Replace_Arms',
                             'Set_Switches',
                             'District Comments',
                             'Above_Grade_Demo',
                             'Dress_Transformer',
                             'Install_Jumpers',
                             'P&C_Work',
                             'Remove_Old_Breakers',
                             'Set_Breakers',
                             'Set_Steel',
                             'Weld_Bus',
                             'Conduit',
                             'Grounding',
                             'Cable Pulling',
                             'Foundations',
                             'Below Grade Demo']]

    ResourceDF.to_csv('ResourcePlan.csv', index=False)



# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    """ This is executed when run from the command line """
    # Setup Logging
    logger = logging.getLogger('root')
    FORMAT = "[%(filename)s:%(lineno)s - %(funcName)20s() ] %(message)s"
    logging.basicConfig(format=FORMAT)
    logger.setLevel(logging.DEBUG)

    main()

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
