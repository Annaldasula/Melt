import streamlit as st
import pandas as pd
import base64
import io
import streamlit as st
import pandas as pd
import numpy as np
import base64
import re
from wordcloud import WordCloud
from PIL import Image
from fuzzywuzzy import fuzz
import matplotlib.pyplot as plt
import gensim
import spacy
import pyLDAvis.gensim_models
from gensim.utils import simple_preprocess
from gensim.models import CoherenceModel
from pprint import pprint
import logging
import warnings
from nltk.corpus import stopwords
import gensim.corpora as corpora
from io import BytesIO
import nltk
import os


# Load data function
def load_data(file):
    if file:
        data = pd.read_excel(file)
        return data
    return None

# Data preprocessing function (You can include your data preprocessing here)

# Function to create separate Excel sheets by Entity
def create_entity_sheets(data, writer):
    # Define a format with text wrap
    wrap_format = writer.book.add_format({'text_wrap': True})

    for Entity in finaldata['Entity'].unique():
        entity_df = finaldata[finaldata['Entity'] == Entity]
        entity_df.to_excel(writer, sheet_name=Entity, index=False)
        worksheet = writer.sheets[Entity]
        worksheet.set_column(1, 4, 48, cell_format=wrap_format)
        # Calculate column widths based on the maximum content length in each column except columns 1 to 4
        max_col_widths = [
            max(len(str(value)) for value in entity_df[column])
            for column in entity_df.columns[5:]  # Exclude columns 1 to 4
        ]

        # Set the column widths dynamically for columns 5 onwards
        for col_num, max_width in enumerate(max_col_widths):
            worksheet.set_column(col_num + 5, col_num + 5, max_width + 2)  # Adding extra padding for readability       
            
            
# Function to save multiple DataFrames in a single Excel sheet
def multiple_dfs(df_list, sheets, file_name, spaces, comments):
    writer = pd.ExcelWriter(file_name, engine='xlsxwriter')
    row = 2
    for dataframe, comment in zip(df_list, comments):
        pd.Series(comment).to_excel(writer, sheet_name=sheets, startrow=row,
                                    startcol=1, index=False, header=False)
        dataframe.to_excel(writer, sheet_name=sheets, startrow=row + 1, startcol=0)
        row = row + len(dataframe.index) + spaces + 2
    writer.close()
     
    
def top_10_dfs(df_list, file_name, comments, top_11_flags):
    writer = pd.ExcelWriter(file_name, engine='xlsxwriter')
    row = 2
    for dataframe, comment, top_11_flag in zip(df_list, comments, top_11_flags):
        if top_11_flag:
            top_df = dataframe.head(50)  # Select the top 11 rows for specific DataFrames
        else:
            top_df = dataframe  # Leave other DataFrames unchanged

        top_df.to_excel(writer, sheet_name="Top 10 Data", startrow=row, index=True)
        row += len(top_df) + 2  # Move the starting row down by len(top_df) + 2 rows

    # Create a "Report" sheet with all the DataFrames
    for dataframe, comment in zip(df_list, comments):
        dataframe.to_excel(writer, sheet_name="Report", startrow=row, index=True, header=True)
        row += len(dataframe) + 2  # Move the starting row down by len(dataframe) + 2 rows

    writer.close()

    
# Streamlit app with a sidebar layout
st.set_page_config(layout="wide")

# Custom CSS for title bar position
title_bar_style = """
    <style>
        .title h1 {
            margin-top: -10px; /* Adjust this value to move the title bar up or down */
        }
    </style>
"""

st.markdown(title_bar_style, unsafe_allow_html=True)

# Modify the paths according to your specific directory
download_path = r"C:\Users\akshay.annaldasula"

st.title("Meltwater Data Insights Dashboard")

# Sidebar for file upload and download options
st.sidebar.title("Upload a file for tables")

# File Upload Section
file = st.sidebar.file_uploader("Upload Data File (Excel or CSV)", type=["xlsx", "csv"])

if file:
    st.sidebar.write("File Uploaded Successfully!")

    # Load data
    data = load_data(file)

    if data is not None:
        # Data Preview Section
#         st.write("## Data Preview")
#         st.write(data)

        # Entity SOV Section
        # st.sidebar.write("## Entity Share of Voice (SOV)")
        # Include your Entity SOV code here

        # Data preprocessing
        data.drop(columns=data.columns[10:], axis=1, inplace=True)
        data = data.rename({'Influencer': 'Journalist'}, axis=1)
        data.drop_duplicates(subset=['Date', 'Entity', 'Headline', 'Publication Name'], keep='first', inplace=True, ignore_index=True)
        data.drop_duplicates(subset=['Date', 'Entity', 'Opening Text', 'Publication Name'], keep='first', inplace=True, ignore_index=True)
        data.drop_duplicates(subset=['Date', 'Entity', 'Hit Sentence', 'Publication Name'], keep='first', inplace=True, ignore_index=True)
        finaldata = data
        En_sov = pd.crosstab(finaldata['Entity'], columns='News Count', values=finaldata['Entity'], aggfunc='count').round(0)
        En_sov.sort_values('News Count', ascending=False)
        En_sov['% '] = ((En_sov['News Count'] / En_sov['News Count'].sum()) * 100).round(2)
        Sov_table = En_sov.sort_values(by='News Count', ascending=False)
        Sov_table.loc['Total'] = Sov_table.sum(numeric_only=True, axis=0)
        Entity_SOV1 = Sov_table.round()

        # st.sidebar.write(Entity_SOV1)

        finaldata['Date'] = pd.to_datetime(finaldata['Date']).dt.normalize()
        sov_dt = pd.crosstab((finaldata['Date'].dt.to_period('M')),finaldata['Entity'],margins = True ,margins_name='Total')

        pub_table = pd.crosstab(finaldata['Publication Name'],finaldata['Entity'])
        pub_table['Total']= pub_table.sum(axis=1)
        pubs_table=pub_table.sort_values('Total',ascending=False).round()
        pubs_table.loc['GrandTotal']= pubs_table.sum(numeric_only=True,axis=0)

        PP = pd.crosstab(finaldata['Publication Name'],finaldata['Publication Type'])
        PP['Total']= PP.sum(axis=1)
        PP_table=PP.sort_values('Total',ascending=False).round()
        PP_table.loc['GrandTotal']= PP_table.sum(numeric_only=True,axis=0)

        PT_Entity = pd.crosstab(finaldata['Publication Type'],finaldata['Entity'])
        PT_Entity['Total']= PT_Entity.sum(axis=1)
        PType_Entity=PT_Entity.sort_values('Total',ascending=False).round()
        PType_Entity.loc['GrandTotal']= PType_Entity.sum(numeric_only=True,axis=0)

        ppe = pd.crosstab(columns=finaldata['Entity'],index=[finaldata["Publication Type"],finaldata["Publication Name"]],margins=True,margins_name='Total')
        ppe1 = ppe.reset_index()
        ppe1.set_index("Publication Type", inplace = True)

        finaldata['Journalist']=finaldata['Journalist'].str.split(',')
        finaldata = finaldata.explode('Journalist')
        jr_tab=pd.crosstab(finaldata['Journalist'],finaldata['Entity'])
        jr_tab = jr_tab.reset_index(level=0)
        newdata = finaldata[['Journalist','Publication Name']]
        Journalist_Table = pd.merge(jr_tab, newdata, how='inner',
                  left_on=['Journalist'],
                  right_on=['Journalist'])

        Journalist_Table.drop_duplicates(subset=['Journalist'], keep='first', inplace=True, ignore_index=True)
        valid_columns = Journalist_Table.select_dtypes(include='number').columns
        Journalist_Table['Total'] = Journalist_Table[valid_columns].sum(axis=1)

        Journalist_Table.sort_values('Total',ascending=False)
        Jour_table=Journalist_Table.sort_values('Total',ascending=False).round()
        bn_row = Jour_table.loc[Jour_table['Journalist'] == 'Bureau News']
        Jour_table = Jour_table[Jour_table['Journalist'] != 'Bureau News']
        Jour_table = pd.concat([Jour_table, bn_row], ignore_index=True)
        Jour_table.loc['GrandTotal'] = Jour_table.sum(numeric_only=True, axis=0)
        Jour_table.insert(1, 'Publication Name', Jour_table.pop('Publication Name'))
        
        # Remove square brackets and single quotes from the 'Journalist' column
        #data['Journalist'] = data['Journalist'].str.strip("[]'")
        
        # Remove square brackets and single quotes from the 'Journalist' column
        data['Journalist'] = data['Journalist'].str.replace(r"^\['(.+)'\]$", r"\1", regex=True)
        
        # Define a function to classify news as "Exclusive" or "Not Exclusive" for the current entity
        def classify_exclusivity(row):
                
            entity_name = finaldata['Entity'].iloc[0]  # Get the entity name for the current sheet
            # Check if the entity name is mentioned in either 'Headline' or 'Similar_Headline'
            if entity_name.lower() in row['Headline'].lower() or entity_name.lower() in row['Headline'].lower():
                
                return "Exclusive"
            else:
                
                return "Not Exclusive"
                    
        # Apply the classify_exclusivity function to each row in the current entity's data
        finaldata['Exclusivity'] = finaldata.apply(classify_exclusivity, axis=1) 
        
        # Define a dictionary of keywords for each entity
        entity_keywords = {
                        'Amazon': ['Amazon','Amazons','amazon'],
#                           'LTTS': ['LTTS', 'ltts'],
#                           'KPIT': ['KPIT', 'kpit'],
#                          'Cyient': ['Cyient', 'cyient'], 
            }
            
        # Define a function to qualify entity based on keyword matching
        def qualify_entity(row):    
            
            entity_name = row['Entity']
            text = row['Headline']   
                
            if entity_name in entity_keywords:
                keywords = entity_keywords[entity_name]
                # Check if at least one keyword appears in the text
                if any(keyword in text for keyword in keywords):
                    
                    return "Qualified"
                
            return "Not Qualified"
            
        # Apply the qualify_entity function to each row in the current entity's data
        finaldata['Qualification'] = finaldata.apply(qualify_entity, axis=1)
        
        # Define a dictionary to map predefined words to topics
        topic_mapping = {
              'Merger': ['merger', 'merges'],
                
              'Acquire': ['acquire', 'acquisition', 'acquires'],
                
              'Partnership': ['partnership', 'tieup', 'tie-up','mou','ties up','ties-up','joint venture'],
                
               'Business Strategy': ['launch', 'launches', 'launched', 'announces','announced', 'announcement','IPO','campaign','launch','launches','ipo','sales','sells','introduces','announces','introduce','introduced','unveil',
                                    'unveils','unveiled','rebrands','changes name','bags','lays foundation','hikes','revises','brand ambassador','enters','ambassador','signs','onboards','stake','stakes','to induct','forays','deal'],
                
               'Investment and Funding': ['invests', 'investment','invested','funding', 'raises','invest','secures'],
                
              'Employee Engagement': ['layoff', 'lay-off', 'laid off', 'hire', 'hiring','hired','appointment','re-appoints','reappoints','steps down','resigns','resigned','new chairman','new ceo','layoffs','lay offs'],
                
              'Financial Performence': ['quarterly results', 'profit', 'losses', 'revenue','q1','q2','q3','q4'],
                
               'Business Expansion': ['expansion', 'expands', 'inaugration', 'inaugrates','to open','opens','setup','set up','to expand','inaugurates'], 
                
               'Leadership': ['in conversation', 'speaking to', 'speaking with','ceo'], 
                
               'Stock Related': ['buy', 'target', 'stock','shares' ,'stocks','trade spotlight','short call','nse'], 
                
                'Awards & Recognition': ['award', 'awards'],
                
                'Legal & Regulatory': ['penalty', 'fraud','scam','illegal'],
            
            'Sale - Offers - Discounts' : ['sale','offers','discount','discounts','discounted']
        }
            
        # Define a function to classify headlines into topics
        def classify_topic(headline):
            
            lowercase_headline = headline.lower()
            for topic, words in topic_mapping.items():
                for word in words:
                    if word in lowercase_headline:
                        return topic
            return 'Other'  # If none of the predefined words are found, assign 'Other'
            
                      
        # Apply the classify_topic function to each row in the dataframe
        finaldata['Topic'] = finaldata['Headline'].apply(classify_topic)
        
        


        dfs = [Entity_SOV1, sov_dt, pubs_table, Jour_table, PType_Entity, PP_table, ppe1]

        comments = ['SOV Table', 'Month-on-Month Table', 'Publication Table', 'Journalist Table','Pub Type and Pub Name Table', 'Pub Type and Entity Table', 'PubType PubName and Entity Table']

        # Sidebar for download options
        st.sidebar.write("## Download Options")
        download_formats = st.sidebar.selectbox("Select format:", ["Excel", "CSV", "Excel (Entity Sheets)"])
        file_name_data = st.sidebar.text_input("Enter file name for all DataFrames", "entitydata.xlsx")

        if st.sidebar.button("Download Data"):
            if download_formats == "Excel":
                # Create a link to download the Excel file for data
                excel_path = os.path.join(download_path, "data.xlsx")
                with pd.ExcelWriter(excel_path, engine="xlsxwriter", mode="xlsx") as writer:
                    data.to_excel(writer, index=False)

                st.sidebar.write(f"Excel file saved at {excel_path}")
#                 excel_io_data = io.BytesIO()
#                 with pd.ExcelWriter(excel_io_data, engine="xlsxwriter", mode="xlsx") as writer:
#                     data.to_excel(writer, index=False)
#                 excel_io_data.seek(0)
#                 b64_data = base64.b64encode(excel_io_data.read()).decode()
#                 href_data = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64_data}" download="data.xlsx">Download Data Excel</a>'
#                 st.sidebar.markdown(href_data, unsafe_allow_html=True)

            elif download_formats == "CSV":
                # Create a link to download the CSV file for data
#                 csv_io_data = io.StringIO()
                csv_path = os.path.join(download_path, "data.csv")
                data.to_csv(csv_path, index=False)
                st.sidebar.write(f"CSV file saved at {csv_path}")
                
#                 data.to_csv(csv_io_data, index=False)
#                 csv_io_data.seek(0)
#                 b64_data = base64.b64encode(csv_io_data.read().encode()).decode()
#                 href_data = f'<a href="data:text/csv;base64,{b64_data}" download="data.csv">Download Data CSV</a>'
#                 st.sidebar.markdown(href_data, unsafe_allow_html=True)

            elif download_formats == "Excel (Entity Sheets)":
                # Create a link to download separate Excel sheets by Entity
#                 excel_io_sheets = io.BytesIO()
                excel_path_sheets = os.path.join(download_path, file_name_data)
                with pd.ExcelWriter(excel_path_sheets, mode="w", date_format='yyyy-mm-dd', datetime_format='yyyy-mm-dd') as writer:
                    create_entity_sheets(data, writer)

                st.sidebar.write(f"Excel sheets saved at {excel_path_sheets}")
#                 with pd.ExcelWriter(excel_io_sheets, engine="xlsxwriter", mode="xlsx" , date_format='yyyy-mm-dd', datetime_format='yyyy-mm-dd') as writer:
#                     create_entity_sheets(data, writer)
#                 excel_io_sheets.seek(0)
#                 b64_sheets = base64.b64encode(excel_io_sheets.read()).decode()
#                 href_sheets = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64_sheets}" download="{file_name_data}">Download Entity Sheets Excel</a>'
#                 st.sidebar.markdown(href_sheets, unsafe_allow_html=True)

        # Download selected DataFrame
        st.sidebar.write("## Download Selected DataFrame")
        
                # Create a dropdown to select the DataFrame to download
        dataframes_to_download = {
            "Entity_SOV1": Entity_SOV1,
            "Data": data,
            "Finaldata": finaldata,
            "Month-on-Month":sov_dt,
            "Publication Table":pubs_table,
            "Journalist Table":Jour_table,
            "Publication Type and Name Table":PP_table,
            "Publication Type Table with Entity":PType_Entity,
            "Publication type,Publication Name and Entity Table":ppe1,
            "Entity-wise Sheets": finaldata  # Add this option to download entity-wise sheets
        }
        
        selected_dataframe = st.sidebar.selectbox("Select DataFrame:", list(dataframes_to_download.keys()))

        if st.sidebar.button("Download Selected DataFrame"):
            if selected_dataframe in dataframes_to_download:
                # Create a link to download the selected DataFrame in Excel
                selected_df = dataframes_to_download[selected_dataframe]
                excel_io_selected = io.BytesIO()
                with pd.ExcelWriter(excel_io_selected, engine="xlsxwriter", mode="xlsx") as writer:
                    selected_df.to_excel(writer, index=True)
                excel_io_selected.seek(0)
                b64_selected = base64.b64encode(excel_io_selected.read()).decode()
                href_selected = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64_selected}" download="{selected_dataframe}.xlsx">Download {selected_dataframe} Excel</a>'
                st.sidebar.markdown(href_selected, unsafe_allow_html=True)
                 
        # Download All DataFrames as a Single Excel Sheet
        st.sidebar.write("## Download All DataFrames as a Single Excel Sheet")
        file_name_all = st.sidebar.text_input("Enter file name for all DataFrames", "all_dataframes.xlsx")
#         download_options = st.sidebar.selectbox("Select Download Option:", [ "Complete Dataframes"])
        
        if st.sidebar.button("Download All DataFrames"):
            # List of DataFrames to save
            dfs = [Entity_SOV1, sov_dt, pubs_table, Jour_table, PType_Entity, PP_table, ppe1]
            comments = ['SOV Table', 'Month-on-Month Table', 'Publication Table', 'Journalist Table',
                        'Pub Type and Entity Table', 'Pub Type and Pub Name Table',
                        'PubType PubName and Entity Table']
            
            excel_path_all = os.path.join(download_path, file_name_all)
            multiple_dfs(dfs, 'Tables', excel_path_all, 2, comments)
            st.sidebar.write(f"All DataFrames saved at {excel_path_all}")
            
#             # Create a link to download all DataFrames as a single Excel sheet with separation
#             excel_io_all = io.BytesIO()
#             multiple_dfs(dfs, 'Tables', excel_io_all, 2, comments)
#             excel_io_all.seek(0)
#             b64_all = base64.b64encode(excel_io_all.read()).decode()
#             href_all = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64_all}" download="{file_name_all}">Download All DataFrames Excel</a>'
#             st.sidebar.markdown(href_all, unsafe_allow_html=True)
            
            
        # Download Top 10 DataFrames as a Single Excel Sheet
        st.sidebar.write("## Download Top N DataFrames as a Single Excel Sheet")
        file_name_topn = st.sidebar.text_input("Enter file name for all DataFrames", "top_dataframes.xlsx")
        # Slider to select the range of dataframes
        selected_range = st.sidebar.slider("Select start range:", 10, 50, 10)
        
        if st.sidebar.button("Download Top DataFrames"):
            # List of DataFrames to save
            selected_dfs = [Entity_SOV1, sov_dt, pubs_table, Jour_table, PType_Entity, PP_table, ppe1]
            comments_selected = ['SOV Table', 'Month-on-Month Table', 'Publication Table', 'Journalist Table',
                        'Pub Type and Entity Table', 'Pub Type and Pub Name Table',
                        'PubType PubName and Entity Table']
            top_n_flags = [False, False, True, True, True, True, True]
            
            # Create a link to download all DataFrames as a single Excel sheet with two sheets
            selected_dfs = [df.head(selected_range) for df in selected_dfs]
            comments_selected = comments_selected[:selected_range]
            top_n_flags = top_n_flags[:selected_range]

            excel_path_topn = os.path.join(download_path, file_name_topn)
            top_10_dfs(selected_dfs, excel_path_topn, comments_selected, top_n_flags)
            st.sidebar.write(f"Selected DataFrames saved at {excel_path_topn}")
            
#             excel_io_all = io.BytesIO()
#             top_10_dfs(dfs1, excel_io_all, commentss, top_11_flags)  # Save the top 10 rows in the first sheet                      
#             excel_io_all.seek(0)
#             b64_all = base64.b64encode(excel_io_all.read()).decode()
#             href_all = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64_all}" download="file_name_top10">Download Top10 DataFrames Excel</a>'
#             st.sidebar.markdown(href_all, unsafe_allow_html=True)
            
    else:
        st.sidebar.write("Please upload a file.")
        
            # Preview selected DataFrame in the main content area
    st.write("## Preview Selected DataFrame")
    selected_dataframe = st.selectbox("Select DataFrame to Preview:", list(dataframes_to_download.keys()))
    st.dataframe(dataframes_to_download[selected_dataframe])

# Add more sections, charts, and widgets as needed


import streamlit as st
import pandas as pd
import numpy as np
import base64
import re
from wordcloud import WordCloud
from PIL import Image
from fuzzywuzzy import fuzz
import matplotlib.pyplot as plt
import gensim
import spacy
import pyLDAvis.gensim_models
from gensim.utils import simple_preprocess
from gensim.models import CoherenceModel
from pprint import pprint
import logging
import warnings
from nltk.corpus import stopwords
import gensim.corpora as corpora
from io import BytesIO
import nltk

# Download NLTK stopwords
nltk.download('stopwords')

# Set up logging
logging.basicConfig(format='%(asctime)s : %(levelname)s : %(message)s', level=logging.ERROR)

# Ignore warnings
warnings.filterwarnings("ignore", category=DeprecationWarning)

# Initialize Spacy 'en' model
nlp = spacy.load('en_core_web_sm', disable=['parser', 'ner'])


# Define a function to clean the text
def clean(text):
    text = text.lower()
    text = re.sub('[^A-Za-z]+', ' ', text)
    text = re.sub('[,\.!?]', ' ', text)
    return text

# Streamlit app with a sidebar layout
# st.set_page_config(layout="wide")

# Custom CSS for title bar position
title_bar_style = """
    <style>
        .title h1 {
            margin-top: -10px; /* Adjust this value to move the title bar up or down */
        }
    </style>
"""

st.markdown(title_bar_style, unsafe_allow_html=True)

st.title("SimilarNews , Wordcloud and Topic Explorer")

# Sidebar for file upload
st.sidebar.title("Upload a file for Data Analysis")

file = st.sidebar.file_uploader("Upload Excel File", type=["xlsx"])

if file:
    st.sidebar.write("File Uploaded Successfully!")

    # Importing Dataset
    data = pd.read_excel(file)
    
    data['Text'] = (data['Headline'].astype(str) + data['Opening Text'].astype(str) + data['Hit Sentence'].astype(str))
    data.drop_duplicates(subset=['Date', 'Entity', 'Headline', 'Publication Name'], keep='first', inplace=True, ignore_index=True)
    data.drop_duplicates(subset=['Date', 'Entity', 'Opening Text', 'Publication Name'], keep='first', inplace=True, ignore_index=True)
    data.drop_duplicates(subset=['Date', 'Entity', 'Hit Sentence', 'Publication Name'], keep='first', inplace=True, ignore_index=True)

    # Define a function to clean the text
    def clean(text):
        text = text.lower()
        text = re.sub('[^A-Za-z]+', ' ', text)
        text = re.sub('[,\.!?]', ' ', text)
        return text

    # Cleaning the text in the Headline column
    data['Cleaned_Headline'] = data['Headline'].apply(clean)

    # Define a function to clean the text
    def cleaned(text):
        # Removes all special characters and numericals leaving the alphabets
        text = re.sub('[^A-Za-z]+', ' ', text)
        text = re.sub(r'[[0-9]*]', ' ', text)
        text = re.sub('[,\.!?]', ' ', text)
        text = re.sub('[\\n]', ' ', text)
        text = re.sub(r'\b\w{1,3}\b', '', text)
        # removing apostrophes
        text = re.sub("'s", '', str(text))
        # removing hyphens
        text = re.sub("-", ' ', str(text))
        text = re.sub("— ", '', str(text))
        # removing quotation marks
        text = re.sub('\"', '', str(text))
        # removing any reference to outside text
        text = re.sub("[\(\[].*?[\)\]]", "", str(text))
        return text

    # Cleaning the text in the review column
    data['Text'] = data['Text'].apply(cleaned)
    data.head()    
    
    st.sidebar.header("Select An Analysis you want to Work On")
    analysis_option = st.sidebar.selectbox(" ", ["Similarity News", "Word Cloud" ,"LDA"])

    # Define the 'entities' variable outside of the conditional blocks
    entities = list(data['Entity'].unique())

    # Define an empty 'wordclouds' dictionary
    wordclouds = {}

    if analysis_option == "Similarity News":
        st.header("Similar News")
        st.sidebar.subheader("Similarity News Parameters")
        # Place your parameters for Similarity News here

        # Create a new workbook to store the updated sheets
        updated_workbook = pd.ExcelWriter('Similar_News_Grouped.xlsx', engine='xlsxwriter')

        # Iterate over unique entities
        for entity in entities:
            # Filter data for the current entity
            entity_data = data[data['Entity'] == entity].copy()

            # for each unique value in Cleaned_Headline within the entity
            for headline in entity_data['Cleaned_Headline'].unique():
                # Compute Levenshtein distance and set to True if >= a limit
                entity_data[headline] = entity_data['Cleaned_Headline'].apply(lambda x: fuzz.ratio(x, headline) >= 70)

                # Set a name for the group (the shortest headline)
                m = np.min(entity_data[entity_data[headline] == True]['Cleaned_Headline'])

                # Assign the group
                entity_data.loc[entity_data['Cleaned_Headline'] == headline, 'Similar_Headline'] = m

            # Drop unnecessary columns
            entity_data.drop(entity_data.columns[36:], axis=1, inplace=True)

            # Sort the dataframe based on the 'Similar_Headline' column
            entity_data.sort_values('Similar_Headline', ascending=True, inplace=True)

            headline_index = entity_data.columns.get_loc('Similar_Headline')

            entity_data = entity_data.iloc[:, :headline_index + 1]

            column_to_delete = entity_data.columns[entity_data.columns.get_loc('Similar_Headline') - 1]  # Get the column name before 'group'
            entity_data = entity_data.drop(column_to_delete, axis=1)
            
            # Define a function to classify news as "Exclusive" or "Not Exclusive" for the current entity
            def classify_exclusivity(row):
                
                entity_name = entity_data['Entity'].iloc[0]  # Get the entity name for the current sheet
                # Check if the entity name is mentioned in either 'Headline' or 'Similar_Headline'
                if entity_name.lower() in row['Headline'].lower() or entity_name.lower() in row['Similar_Headline'].lower():                
                    return "Exclusive"
                else:
                    return "Not Exclusive"
                    
            # Apply the classify_exclusivity function to each row in the current entity's data
            entity_data['Exclusivity'] = entity_data.apply(classify_exclusivity, axis=1)    
            
            
            # Define a dictionary of keywords for each entity
            entity_keywords = {
                        'Nothing Tech': ['Nothing','nothing'],
#                         'Asian Paints': ['asian', 'keyword2', 'keyword3'],
            }
            
            # Define a function to qualify entity based on keyword matching
            def qualify_entity(row):                
                entity_name = row['Entity']
                text = row['Headline']   
                
                if entity_name in entity_keywords:
                    keywords = entity_keywords[entity_name]
                    # Check if at least one keyword appears in the text
                    if any(keyword in text for keyword in keywords):
                        return "Qualified"
                
                return "Not Qualified"
            
            # Apply the qualify_entity function to each row in the current entity's data
            entity_data['Qualification'] = entity_data.apply(qualify_entity, axis=1)
            
            # Define a dictionary to map predefined words to topics
            topic_mapping = {
              'Merger': ['merger', 'merges'],
                
              'Acquire': ['acquire', 'acquisition', 'acquires'],
                
              'Partnership': ['partnership', 'tieup', 'tie-up','mou','ties up','ties-up','joint venture'],
                
               'Business Strategy': ['launch', 'launches', 'launched', 'announces','announced', 'announcement','IPO','campaign','launch','launches','ipo','sales','sells','introduces','announces','introduce','introduced','unveil',
                                    'unveils','unveiled','rebrands','changes name','bags','lays foundation','hikes','revises','brand ambassador','enters','ambassador','signs','onboards','stake','stakes','to induct','forays','deal'],
                
               'Investment and Funding': ['invests', 'investment','invested','funding', 'raises','invest','secures'],
                
              'Employee Engagement': ['layoff', 'lay-off', 'laid off', 'hire', 'hiring','hired','appointment','re-appoints','reappoints','steps down','resigns','resigned','new chairman','new ceo'],
                
              'Financial Performence': ['quarterly results', 'profit', 'losses', 'revenue','q1','q2','q3','q4'],
                
               'Business Expansion': ['expansion', 'expands', 'inaugration', 'inaugrates','to open','opens','setup','set up','to expand','inaugurates'], 
                
               'Leadership': ['in conversation', 'speaking to', 'speaking with','ceo'], 
                
               'Stock Related': ['buy', 'target', 'stock','shares' ,'stocks','trade spotlight','short call','nse'], 
                
                'Awards & Recognition': ['award', 'awards'],
                
                'Legal & Regulatory': ['penalty', 'fraud','scam','illegal'],
            # Add more topics and corresponding words as needed
}
           
             # Define a function to classify headlines into topics
#             def classify_topic(headline):
#                 for topic, words in topic_mapping.items():
#                     for word in words:
#                         if word.lower() in headline.lower():
#                             return topic
#                 return 'Other'  # If none of the predefined words are found, assign 'Other'
            
            # Define a function to classify headlines into topics
            def classify_topic(headline):
                lowercase_headline = headline.lower()
                for topic, words in topic_mapping.items():
                    for word in words:
                        if word in lowercase_headline:
                            return topic
                return 'Other'  # If none of the predefined words are found, assign 'Other'

            
            
            
            # Apply the classify_topic function to each row in the dataframe
            entity_data['Topic'] = entity_data['Headline'].apply(classify_topic)            
            
            # Save the updated sheet to the new workbook
            entity_data.to_excel(updated_workbook, sheet_name=entity, index=False, startrow=0)

            # Create a Word Cloud for the entity based on the Headline column
            wordcloud = WordCloud(width=550, height=400, background_color='white').generate(' '.join(entity_data['Cleaned_Headline']))
            wordclouds[entity] = (wordcloud, entity_data)

        # Save the new workbook
        updated_workbook.close()

        # Provide a download link for the grouped data
        st.markdown("### Download Grouped Data")
        st.markdown(
            f"Download the grouped data as an Excel file: [Similar_News_Grouped.xlsx](data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{base64.b64encode(open('Similar_News_Grouped.xlsx', 'rb').read()).decode()})"
        )

        # Provide a download link for the grouped data in CSV format
        data_csv = data.to_csv(index=False)
        st.markdown(
            f"Download the original data as a CSV file: [Original_Data.csv](data:text/csv;base64,{base64.b64encode(data_csv.encode()).decode()})"
        )

        # Load the grouped data from the "Similar_News_Grouped.xlsx" file
        grouped_data = pd.read_excel("Similar_News_Grouped.xlsx", sheet_name=None)

        # Data Preview Section
        st.sidebar.subheader("Data Preview")
        entities = list(wordclouds.keys())

        selected_entities = st.sidebar.multiselect("Select Entities to Preview", entities)

        if selected_entities:
            for entity in selected_entities:
                st.header(f"Preview for Entity: {entity}")
                entity_data = grouped_data[entity]
                st.write(entity_data)

    elif analysis_option == "Word Cloud":
        st.header("WordCloud")
        st.sidebar.subheader("Word Cloud Parameters")
        # Place your parameters for Word Cloud here

        # Generate and display word clouds for selected entities
        st.sidebar.title("Word Clouds")

        wordcloud_entity = st.sidebar.selectbox("Select Entity for Word Cloud", entities)

        # Custom Stop Words Section
        st.sidebar.title("Custom Stop Words")

        custom_stopwords = st.sidebar.text_area("Enter custom stop words (comma-separated)", "")
        custom_stopwords = [word.strip() for word in custom_stopwords.split(',')]

        # Widget to adjust word cloud parameters
        wordcloud_size_height = st.slider("Select Word Cloud Size Height", 100, 1000, 400, step=50, key="wordcloud_height")
        wordcloud_size_width = st.slider("Select Word Cloud Size Width", 100, 1000, 400, step=50, key="wordcloud_width")
        wordcloud_max_words = st.slider("Select Max Number of Words", 10, 500, 50)

        if wordcloud_entity:
            st.header(f"Word Cloud for Entity: {wordcloud_entity}")
            # Generate Word Cloud with custom stop words removed
            cleaned_headlines = ' '.join(data[data['Entity'] == wordcloud_entity]['Text'])

            if custom_stopwords:
                for word in custom_stopwords:
                    cleaned_headlines = cleaned_headlines.replace(word, '')

            wordcloud_image = WordCloud(font_path="D:\Akshay.Annaldasula\OneDrive - Adfactors PR Pvt Ltd\Downloads\Helvetica.ttf",
                                        background_color="white", width=wordcloud_size_width, height=wordcloud_size_height, max_font_size=80, max_words=wordcloud_max_words,
                                        colormap='Set1', contour_color='black', contour_width=2, collocations=False).generate(cleaned_headlines)
            
            
            # Create entity_data for the selected entity
            entity_data = data[data['Entity'] == wordcloud_entity]

            # Resize the word cloud image using PIL
            img = Image.fromarray(np.array(wordcloud_image))
            img = img.resize((wordcloud_size_width, wordcloud_size_height))
            
            # Add the entity to the wordclouds dictionary
            wordclouds[wordcloud_entity] = (wordcloud_image, entity_data)

            # Display the resized word cloud image in Streamlit
            st.image(img, caption=f"Word Cloud for Entity: {wordcloud_entity}")

        # Word Cloud Interaction
        if wordcloud_entity:
            st.header(f"Word Cloud Interaction for Entity: {wordcloud_entity}")
            
            # Debugging statements
            st.write("Entities in entities list:", entities)
            st.write("Keys in wordclouds dictionary:", list(wordclouds.keys()))
            
            # Get the selected entity's word cloud
            entity_wordcloud, entity_data = wordclouds.get(wordcloud_entity, (None, None))  # Use .get() to handle missing keys gracefully
            if entity_wordcloud is None:
                st.warning(f"No word cloud found for '{wordcloud_entity}'")
            else:
                words = list(entity_wordcloud.words_.keys())
            

            # Get the selected entity's word cloud
            entity_wordcloud, entity_data = wordclouds[wordcloud_entity]
            words = list(entity_wordcloud.words_.keys())

            word_frequencies = entity_wordcloud.words_
            words_f = list(word_frequencies.keys())

            # Create a list of tuples containing (word, frequency)
            word_frequency_list = [(word, frequency) for word, frequency in word_frequencies.items()]

            # Add this line to preview the words and their frequencies
            st.write("Words and their frequencies:")
            st.write(word_frequency_list)

            # Filter out bigrams from the list of words
            # individual_words = [word for word in words if ' ' not in word]

            # Create a selectbox for the words in the word cloud
            selected_word = st.selectbox("Select a word from the word cloud", words)

            # Find rows where the selected word appears
            matching_rows = entity_data[entity_data['Headline'].str.contains(selected_word, case=False, na=False)]

            # Display the matching rows or a message if no matches are found
            if not matching_rows.empty:
                st.subheader(f"Matching Rows for '{selected_word}':")
                # Function to highlight the selected word
                def highlight_word(text, word):                
                    return re.sub(f'\\b{word}\\b', f'**{word}**', text, flags=re.IGNORECASE)

                # Apply the highlight_word function to the Cleaned_Headline column
                matching_rows['Headline'] = matching_rows.apply(lambda row: highlight_word(row['Headline'], selected_word), axis=1)

                # Display the formatted dataframe
                st.dataframe(matching_rows)

            else:
                st.warning(f"No matching rows found for '{selected_word}'")
    
    
    elif analysis_option == "LDA":
        st.header("Topic Modelling (LDA)")
        selected_entity = st.sidebar.selectbox("Select Entity for LDA", entities)
        
        # Apply LDA to entity_data['Text']
        # Create entity_data for the selected entity
#         entity_data = data[data['Entity'] == wordcloud_entity]

        num_of_topics = st.sidebar.slider("Number of Topics", min_value=1, max_value=20, value=10)
#         per_word_topics = st.sidebar.checkbox("Per Word Topics",
#                                      help="If True, the model also computes a list of topics, sorted in descending order of most likely topics for each word, along with their phi values multiplied by the feature length (i.e. word count).")
        iterations = st.sidebar.number_input("Iterations", min_value=1, value=50,
                                    help="Maximum number of iterations through the corpus when inferring the topic distribution of a corpus.")
        no_words = st.sidebar.number_input("No of Words", min_value=10, value=30,
                                help="Number of words to be displayed in the LDA graph.")       


        cut_off_percentage = st.sidebar.slider("Topic Cutoff Percentage", min_value=0.0, max_value=1.0, value=0.25, step=0.01)
        
        # Run LDA for the selected entity
        if selected_entity:
            st.header(f"Running LDA for entity: {selected_entity}")
            
            # Filter data for the selected entity
            entity_data = data[data['Entity'] == selected_entity]
#             data_text = entity_data['Text'].apply(cleaned).values.tolist()
        

#           uploaded_file = st.sidebar.file_uploader("Upload Excel File", type=["xlsx"])
            file_name = st.sidebar.text_input("Output File Name", f"{selected_entity} new_topics.xlsx") 
            entity_data['Text'] = entity_data['Text'].apply(cleaned)  # Use your text cleaning function if needed
            st.write(entity_data)
            data1 = (entity_data.Text.values.tolist())
#             st.write(data1)
            stop_words = stopwords.words('english')

            # Tokenize and preprocess text data
        
        
            def tokenize_and_preprocess(data1):
                data_words = list(sent_to_words(data1))
                data_words_nostops = remove_stopwords(data_words)
                data_words_bigrams = make_bigrams(data_words_nostops)
                data_lemmatized = lemmatization(data_words_bigrams, allowed_postags=['NOUN', 'ADJ', 'VERB', 'ADV'])
                return data_lemmatized

            def sent_to_words(sentences):
                for sentence in sentences:
                    yield(gensim.utils.simple_preprocess(str(sentence).encode('utf-8'), deacc=True))
            
            data_words = list(sent_to_words(data1))        

            def remove_stopwords(texts):
                return [[word for word in simple_preprocess(str(doc)) if word not in stop_words] for doc in texts]

            def make_bigrams(texts):
                bigram = gensim.models.Phrases(data_words, min_count=5, threshold=100) # higher threshold fewer phrases.
                trigram = gensim.models.Phrases(bigram[data_words], threshold=100)  

                # Faster way to get a sentence clubbed as a trigram/bigram
                bigram_mod = gensim.models.phrases.Phraser(bigram)
                trigram_mod = gensim.models.phrases.Phraser(trigram)
                return [bigram_mod[doc] for doc in texts]

            def lemmatization(texts, allowed_postags=['NOUN', 'ADJ', 'VERB', 'ADV']):
                texts_out = []
                for sent in texts:
                    doc = nlp(" ".join(sent))
                    texts_out.append([token.lemma_ for token in doc if token.pos_ in allowed_postags])
                return texts_out

            data_lemmatized = tokenize_and_preprocess(data1)

            # Build the LDA model
#           @st.cache_data 
            def build_lda_model(corpus, _id2word, num_topics):
                lda_model = gensim.models.ldamodel.LdaModel(
                corpus=corpus,
                id2word=_id2word,
                num_topics=num_topics,
                random_state=100,
                update_every=1,
                chunksize=100,
                passes=10,
                alpha='auto',
                per_word_topics=True,
                iterations=iterations    
        )
                return lda_model

            # Create Dictionary and Corpus
            id2word = corpora.Dictionary(data_lemmatized)
            texts = data_lemmatized
            corpus = [id2word.doc2bow(text) for text in texts]

            # Build the LDA model
            lda_model = build_lda_model(corpus, id2word, num_of_topics)

            # Visualize the topics using pyLDAvis
#           pyLDAvis.enable_notebook()
            vis = pyLDAvis.gensim_models.prepare(lda_model, corpus, id2word,R=no_words)
#           st.title("Topic Modeling Visualization")
#           st.write(vis, use_container_width=True)

            # Create a list to store the assigned topic numbers for each document
            topic_assignments = []

            # Loop through each document in the corpus and assign a topic number
            for doc in corpus:
                # Get the topic probabilities for the document
                topic_probs = lda_model.get_document_topics(doc)

                # Sort the topic probabilities in descending order by probability score
                topic_probs.sort(key=lambda x: x[1], reverse=True)

                # Filter out topics with a contribution less than the cutoff percentage
                topic_probs = [topic for topic in topic_probs if topic[1] >= cut_off_percentage]

                # Get the topic number with the highest probability score
                top_topic_num = topic_probs[0][0] if topic_probs else -1

                # Append the topic number to the list of topic assignments
                topic_assignments.append(top_topic_num)

            # Add the topic assignments to the DataFrame
            entity_data['Topic'] = topic_assignments
    
            st.markdown(
    """
    <style>
    .stApp {
        width: 100%;
    }
    </style>
    """,
            unsafe_allow_html=True,
)

            # Streamlit UI
            st.header(f"LDA Topic Modeling Results {selected_entity}")
            st.dataframe(entity_data)
            st.header(f"Topic Modeling Visualization {selected_entity}")
            st.components.v1.html(pyLDAvis.prepared_data_to_html(vis),width=1500, height=800)

            # Download button for the modified DataFrame
            st.sidebar.markdown("## Download Results")
            st.sidebar.markdown("Click below to download the modified DataFrame:")
    
            # Save the DataFrame to an in-memory Excel file
            excel_buffer = BytesIO()
            with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                entity_data.to_excel(writer, index=False, sheet_name="Sheet1")
        
            # Convert in-memory Excel file to bytes
            excel_data = excel_buffer.getvalue()

            # Provide a download link for the Excel file
            st.sidebar.download_button(
               label="Download Excel",
               data=excel_data,
               file_name=file_name,
               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)    
  

        else:
            st.write("Please select an entity to run LDA.")
            
            
        # Collect all the documents for each topic
        topic_docs = {}
        for i in range(num_of_topics):
            topic_docs[i] = ' '.join(entity_data[entity_data['Topic'] == i]['Text'].values)
            
        # Add a dropdown for selecting topics from the LDA results
        selected_topic_index = st.sidebar.selectbox("Select a topic", range(num_of_topics))
        
        st.header(f"Word Cloud for Topic: {selected_topic_index}")
        
        # Custom Stop Words Section
        st.sidebar.title("Custom Stop Words")

        custom_stopwords = st.sidebar.text_area("Enter custom stop words (comma-separated)", "")
        custom_stopwords = [word.strip() for word in custom_stopwords.split(',')]
        
        # Widget to adjust word cloud parameters
        wordcloud_size_height = st.sidebar.slider("Select Word Cloud Size Height", 100, 1000, 400, step=50, key="wordcloud_height")
        wordcloud_size_width = st.sidebar.slider("Select Word Cloud Size Width", 100, 1000, 400, step=50, key="wordcloud_width")
        wordcloud_max_words = st.slider("Select Max Number of Words", 10, 500, 100)
        
        # Display the word cloud for the selected topic
        if selected_topic_index is not None:
            selected_topic_text = topic_docs[selected_topic_index]
            wordcloud_image = WordCloud(font_path="D:\Akshay.Annaldasula\OneDrive - Adfactors PR Pvt Ltd\Downloads\Helvetica.ttf", background_color='white',colormap='Set1', contour_color='black', contour_width=2, collocations=False,max_font_size=80,width=wordcloud_size_width, height=wordcloud_size_height,stopwords=custom_stopwords, max_words=wordcloud_max_words).generate(selected_topic_text)
#             plt.figure(figsize=(5, 5))
#             plt.imshow(wordcloud, interpolation='bilinear')
#             plt.axis('off')
#             # Display the word cloud using Streamlit's st.pyplot() method
#             fig, ax = plt.subplots()
#             ax.imshow(wordcloud, interpolation='bilinear')
#             ax.axis('off')
#             st.pyplot(fig)
            # Resize the word cloud image using PIL
            img = Image.fromarray(np.array(wordcloud_image))
            img = img.resize((wordcloud_size_width, wordcloud_size_height))
            # Display the resized word cloud image in Streamlit
            st.image(img)
            
        # Iterate over unique topics
        # Assuming 'entity_data' DataFrame already has a column named 'Topics'
        topics = entity_data['Topic'].unique()
        
        for topic in topics:
            # Filter data for the current topic
            topic_data = entity_data[entity_data['Topic'] == topic].copy()

            # for each unique value in Cleaned_Headline within the topic
            for headline in topic_data['Cleaned_Headline'].unique():
                # Compute Levenshtein distance and set to True if >= a limit
                topic_data[headline] = topic_data['Cleaned_Headline'].apply(lambda x: fuzz.ratio(x, headline) >= 70)

                # Set a name for the group (the shortest headline)
                m = np.min(topic_data[topic_data[headline] == True]['Cleaned_Headline'])

                # Assign the group
                topic_data.loc[topic_data['Cleaned_Headline'] == headline, 'Similar_Headline'] = m

            # Drop unnecessary columns
            topic_data.drop(topic_data.columns[36:], axis=1, inplace=True)

            # Sort the dataframe based on the 'Similar_Headline' column
            topic_data.sort_values('Similar_Headline', ascending=True, inplace=True)

            headline_index = topic_data.columns.get_loc('Similar_Headline')

            topic_data = topic_data.iloc[:, :headline_index + 1]

            column_to_delete = topic_data.columns[topic_data.columns.get_loc('Similar_Headline') - 1]  # Get the column name before 'group'
            topic_data = topic_data.drop(column_to_delete, axis=1)

            # Display a preview of the data for the current topic
            if selected_topic_index == topic:
                st.subheader(f"Preview of Similar News for Topic {topic}")
                st.write(topic_data)      
                
                
                # Add a button to download the data as an Excel file
                if st.sidebar.button('Download Topics Data as Excel'):
                    file_name = f"topic_{topic}.xlsx"
#                     excel_data = BytesIO()
#                     topic_data.to_excel(file_name, index=False)
#                     st.markdown(get_download_link(file_name), unsafe_allow_html=True)
#                     # Create a link to download the Excel file for data
                    excel_path = os.path.join(download_path, file_name)
                    topic_data.to_excel(excel_path, index=False)
                    st.sidebar.write(f"Excel file saved at {excel_path}")
                    
                    
#             def get_download_link(file_path):
#                 with open(file_path, 'rb') as file:
#                     file_content = file.read()
#                 base64_encoded = base64.b64encode(file_content).decode()
#                 download_link = f'<a href="data:application/octet-stream;base64,{base64_encoded}" download="{os.path.basename(file_path)}">Download file</a>'
#                 return download_link        







    
