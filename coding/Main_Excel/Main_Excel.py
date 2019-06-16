import xlwings as xw
import pandas as pd


def run1():
    wb = xw.Book.caller()
    # wb = xw.Book('I:\\github_repos\\My_Freelancing_Projects\\work-1\\coding\\Main_Excel\\Main_Excel.xlsm')
    # wb.sheets[0].range("A1").value = "Hello xlwings!"     # test code

    #-----------------------------------------------------------------------------------------------------------
    sht_1 = wb.sheets['Sample-1']   # sheet for 'Sample-1'
    # sht_run = wb.sheets['RUN']

    #------------------------------------------------------------------------------------------------------------
    # Inputs
    excel_file_directory = 'I:\\github_repos\\My_Freelancing_Projects\\work-1\\coding\\Main_Excel\\data\\sample_1_PEX_Our_Food_Nutritionals_spring-converted.xlsx'
    max_sht_num = 11
    # dish_types = []       # if not constant, then fetch the index cell value by looping along the dataframes
    dish_types = ['Starters', 'Bases', 'Romana Pizzas and Calabrese', 
        'Classic Pizzas', 'Leggera Pizzas', 'Salads No Dressings', 
        'Salads With Dressings and Dough Sticks', 'Al Forno', 'Desserts', 
        'Dolcetti', 'Piccolo']

    #---------------------------------------------------------------------------------------------------------------------------------------------
    excel_file = pd.ExcelFile(excel_file_directory)     # Excel file for sample_1_PEX_Our_Food_Nutritionals_spring
    
    df_sample1 = []     # Initialize the list
    columns_header_sample1 = ["Energy kcal", "Energy kJ", "Fat g", "Saturates g", "Carbs g", "Sugars g", "Fibre g", "Protein g", "Salt g", 
                            "Energy kcal", "Energy kJ", "Fat g", "Saturates g", "Carbs g", "Sugars g", "Fibre g", "Protein g", "Salt g"]
    # per_serving_str = 'PER SERVING'
    # per_100g_str = 'PER 100 G'
    # columns_per_serving = [per_serving_str + x for x in columns_header_sam1]
    # columns_per_100g = [per_100g_str + y for y in columns_header_sam1]
    # columns_sample1 = list(set(columns_per_serving + columns_per_100g))
    # ===================================================================================================================
    # M-1: Individual
    # df1 = excel_file.parse('Table 1', skiprows=1)
    # df1 = df1.dropna(axis=1)
    # df1 = df1.set_index(dish_types[0])
    # df1.index.name = None
    # # df1.rename(columns= columns_header_sam1, inplace=True)
    # df_sample1 = [df1, df2, df3]
    # ===================================================================================================================
    # M-2: Looping
    for x in range(1, max_sht_num+1):   # x: 0 to max_sht_num (i.e. 11)
        df = excel_file.parse('Table ' + str(x), skiprows= 1)
        df = df.drop(['Unnamed: 10'], axis= 1)
        df = df.dropna()
        df = df.set_index(dish_types[x-1])
        df.index.name = None
        df_sample1.append(df)

    df_sample1 = pd.concat(df_sample1, keys= dish_types)
    df_sample1.index.name = 'MENU ITEMS'
    df_sample1.columns = columns_header_sample1
    # ----------------------------------------------------------------------------------------------------------------------
    # Displaying the data
    sht_1.clear()   # Clear the content and formatting before displaying the data
    sht_1.range('A1').value = df_sample1
    sht_1.autofit('c')    # autofit the column size w.r.t the value inside
    sht_1.autofit('r')    # autofit the row size w.r.t the value inside

    # df.columns = ['MENU ITEM', 'Celery', 'Cereals with Gluten', 'Crustaceans', 'Egg', 'Fish', 'Lupin', 'Milk', 'Molluscs', 
    #             'Mustard', 'Nuts', 'Peanuts', 'Sesame', 'Soybeans', 'Sulphur Dioxide/ Sulphites', 'Garlic', 'Onions', 'Vegans', 'Vegetarians', '-']
    

# if __name__ == '__main__':
#     run1()

# References
# - https://www.datacamp.com/community/tutorials/joining-dataframes-pandas
# - https://pandas.pydata.org/pandas-docs/stable/user_guide/merging.html