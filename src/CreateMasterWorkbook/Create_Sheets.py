# import keyword
import os
import re
import pandas as pd
from pathlib import Path

# Here goes any needed module and function 
# from config import data_info_original
from data_cleaning import normalize_text, remove_duplicates, handle_missing_values
from consolidation import read_and_prepare_csv



class Create_Sheets(object):
    '''Given the proper directories and files create a workbook containing the
    necessary spreadsheets for Aldo's words project.'''
    def __init__(self, folder_paths, file_names):
        '''Create the object saving the folder and file dictionary.'''
        self.folder_paths = folder_paths
        self.file_names = file_names
        

    def check_member(self, member):
        '''See if a member has been created for this class.'''
        return hasattr(self, member)

    
    def consolidate_data(self):
        '''Concatinate the contents of each CSV file putting them in a dataframe'''
        folder_path_original = Path(self.folder_paths['original'])
        
        # Create a list of dataframes from csv files foun in the folder_path_original
        dataframes = []
        for file_path in folder_path_original.glob('*.csv'):
            df_prepared = read_and_prepare_csv(file_path)
            if df_prepared is not None:
                dataframes.append(df_prepared)
            else:
                print(f"Skipping file due to read errors: {file_path}")

        if not dataframes:
            raise("No valid CSV files were found or successfully read.")

        # Combine the dataframes in the list
        self.df_original = pd.concat(dataframes, ignore_index=True)
        self.df_original.sort_values(by='Volume', ascending=False, inplace=True)

    def print_originals(self):
        '''Print out the sentences and associated data that were found'''
        print(f'Original sentences:\n {self.df_original}\n')
         
    def words_sheet(self):
        '''Clean the original sentences, break them into words and save the
        words to a dataflrame then create a dataframe with the unique words
        and the count of times each word appears in the word dataframe.'''
        self.consolidate_data()

        if hasattr(self, 'df_original'):
            # Applying cleaning functions
            # Change to lower case and get rid of non alphnumeric characters and punctuation
            self.df_clean = normalize_text(self.df_original , ['Keyword'])  # Assumes 'Keyword' is the column to be cleaned
            self.df_clean = remove_duplicates(self.df_clean , ['Keyword'])
            self.df_clean = handle_missing_values(self.df_clean , ['Keyword'])
            
            # Split into words count and delete dupes 
            # Reset the index to prevent index numbers in the output
            self.df_clean = self.df_clean.reset_index(drop=True)
            
            # Split the text strings into words and flatten the list
            sentences = self.df_clean['Keyword'].tolist()
            words_list = []
            words = []
            for sentence in sentences:
                words = sentence.split()
                words_list.extend(words)
            # Create a new DataFrame with the words in a single column
            self.df_words = pd.DataFrame(words_list, columns=['Keyword'])  
            # Count the dupes and get rid of them
            series_words = self.df_words.value_counts()
            self.df_counts = series_words.reset_index(name='Count')
        else:
            raise(ValueError('Original dataframe could not be created.'))
        
    def print_clean_sentences(self):
        '''Print out the sentences after they are scrubbed'''
        print(f'Clean sentences: \n{self.df_clean}\n')
       
    def print_words(self):
        '''Print out the words from their sentences.'''
        print(f'Words with counts: \n{self.df_words}\n')
      
    def print_counts(self):
        '''Print out the words and there counts.'''
        print(f'Words with counts: \n{self.df_words}\n')
          
    def excludes_sheet(self):
        '''Read the always and also CSV files and combine them into one sheet
        with two columns.'''
        # Read data from the always and also csv files into dataframes
        # Build the fully qualified file names for the exclude files
        always_file = os.path.join(self.folder_paths['reference'], self.file_names['exclude_always'])
        also_file = os.path.join(self.folder_paths['reference'], self.file_names['exclude_also'])

        # Read data from the aways CSV file into a DataFrame
        self.df_always = pd.read_csv(always_file)

        # Read data from the also CSV file into a DataFrame
        self.df_also = pd.read_csv(also_file)

        # Merge data from both DataFrames
        self.df_excludes = pd.concat([self.df_always.rename(columns={'col_exclude_always': 'Exclude'}), 
                               self.df_also.rename(columns={'col_exclude_also': 'Exclude'})], ignore_index=True)
        self.df_always.rename(columns={'Exclude': 'col_exclude_always'}, inplace=True)
        self.df_also.rename(columns={'Exclude': 'col_exclude_also'}, inplace=True)

    def print_excludes(self):
        ''' Print out the counts of the columns, their combined count and the dataframe.'''
        print(f'Always exclude word counts: {self.df_always['col_exclude_always'].count()}.\n',
              f'Also exclude word counts: {self.df_also['col_exclude_also'].count()}.\n',
              f'Combined count: {self.df_excludes['exclude'].count()}\n')
        print(f'All Excluded words: \n{self.df_excludes}')
    
    def remove_excludes(self):
        '''Taking the words in the counts dataframe, remove any that are in the
        excluded words dataframe and put them in the irrelevant dataframe and
        dump the rest into the relevant dataframe'''
        # Since we are creating the sheets from scratch there will be no
        # new words to add to the exclude also list.  Also the relevant
        # and irrelevant list will be empty to start with so we look at
        # the counts list to see if any removals are necessary.

        # Create a set of all excluded words.
        excludes = set(self.df_excludes['Exclude'].tolist())
        
        self.df_irrelevant = self.df_counts[self.df_counts['Keyword'].isin(excludes)].copy(deep=True)
        self.df_relevant = self.df_counts[~self.df_counts['Keyword'].isin(excludes)].copy(deep=True)
        
        print(f"{self.df_irrelevant['Keyword'].count()} excluded words removed from the Relevant sheet.")

    def create_excel(self, file_name):
        '''Create a two column dataframe for the Excludes sheet the write each
        dataframe out to its sheet in the workbook'''
        # Remove any excluded words from df_words
        self.remove_excludes()
        
        # Combine the exclude dataframes into one with two columns to create
        # a single sheet in excel
        
        # The next three lines will add NaNs to the columns to ensure they are the
        # same length.  If they are not they will not join properly or at all.
        max_len = max(len(self.df_always), len(self.df_also))
        self.df_always = self.df_always.reindex(range(max_len))
        self.df_also = self.df_also.reindex(range(max_len))

        df_exclude = pd.concat([self.df_always, self.df_also], axis=1)
        df_exclude.reset_index(drop=True, inplace=True)
        

        # Create a Pandas Excel writer using the desired file name
        with pd.ExcelWriter(file_name) as writer:
            # Write each DataFrame to a separate sheet
            self.df_original.sort_values(by='Volume', ascending=False, inplace=True)
            self.df_original.to_excel(writer, sheet_name='Original', index=False)

            self.df_clean.sort_values(by='Volume', ascending=False, inplace=True)
            self.df_clean.to_excel(writer, sheet_name='Clean', index=False)
            
            self.df_words.to_excel(writer, sheet_name='Words', index=False)
            
            self.df_counts.sort_values(by='Count', ascending=False, inplace=True)
            self.df_counts.to_excel(writer, sheet_name='Counts', index=False)
            
            self.df_irrelevant.sort_values(by='Count', ascending=False, inplace=True)
            self.df_irrelevant.to_excel(writer, sheet_name='Irrelevant', index=False)            
 
            # Add a column to select words to add to exclude list
            self.df_relevant['Exclude'] = False
            self.df_relevant.sort_values(by='Count', ascending=False, inplace=True)
            self.df_relevant.to_excel(writer, sheet_name='Relevant', index=False)
            
            df_exclude.to_excel(writer, sheet_name='Excludes', index=False)
 
            