# from xmlrpc.client import boolean
import pandas as pd
import numpy as np
# import os

class remove_excludes(object):
    '''Looks for and removes words in the exclude list and moves new excludes.'''
    def __init__(self, file_name):
        '''Generates object and sets target file name.'''
        self.excel_file = file_name
    
    def check_member(self, member):
        '''See if the object contains the given member'''
        return hasattr(self, member)

    def Create_DataFrames(self):
        '''Read the given excel file and converting each sheet to a dataframe.'''
        try:
            # ExcelWriter seems to delete the target workbook when its created
            # so make sure we have all the data for each sheet before calling.
            # Get the original sheet so we can save it back out it is not used.
            self.df_original = pd.read_excel(self.excel_file, sheet_name='Original')
            self.df_clean = pd.read_excel(self.excel_file, sheet_name='Clean')
            self.df_words = pd.read_excel(self.excel_file, sheet_name='Words')
            self.df_counts = pd.read_excel(self.excel_file, sheet_name='Counts')
            self.df_irrelevant = pd.read_excel(self.excel_file, sheet_name='Irrelevant')
            # Keep pandas from returning numbers for booleans
            dtype_dict = {'Exclude': bool}
            self.df_relevant = pd.read_excel(self.excel_file, sheet_name='Relevant', dtype=dtype_dict)            
            self.df_excludes = pd.read_excel(self.excel_file, sheet_name='Excludes')
        except Exception as e:
            print('There was a problem opening the in', self.excel_file, '\nThe error returned was:',e)
        
      

    def create_excludes_set(self):
        '''Combines always and also columns and uses the result to create a set,'''
        # Make sure we've opened the book
        if not self.check_member('df_excludes'):
            self.Create_DataFrames()
            
        always_excludes_list = []
        always_excludes_list = self.df_excludes['col_exclude_always'].tolist()
        # Remove any non string entries from the list like NaNs.
        always_excludes_list = [x for x in always_excludes_list if isinstance(x, str)]

        also_excludes_list = []
        also_excludes_list = self.df_excludes['col_exclude_also'].tolist()
        # Remove any non string entries from the list like NaNs.
        also_excludes_list = [x for x in also_excludes_list if isinstance(x, str)]

        # Concatenate the two Series and convert to set
        always_excludes_list.extend(also_excludes_list)
        all_excludes_set = set(always_excludes_list)
        return all_excludes_set
            
    def move_words_to_exclude_also(self):
        '''Any rows in which the Exclude column is true are moved to the also list'''
        if not self.check_member('df_relevant'):
            self.Create_DataFrames()
        if self.df_relevant['Exclude'].any(bool_only=True, skipna=True):
            # Filter original dataframe to get rows where the column 'Exclude' is True
            # This is a slice of the self.df_relevant data rather than a copy
            self.df_irrelevant = self.df_relevant[self.df_relevant['Exclude']].copy(deep=True)
            self.df_irrelevant.drop(columns=['Exclude'], axis=1, inplace=True)
            # Drop rows in original dataframe corresponding to True values
            self.df_relevant = self.df_relevant[~self.df_relevant['Exclude']]
            num_rec_dropped = len(self.df_irrelevant)
            
            # Make a real copy so that operations on the DataFrame that might affect the
            # silce df_words_to_drop do not effect the information to be added to the
            # df_excludes['col_exclude_also'] column
            df_also_adds = self.df_irrelevant.copy(deep=True)
            # Make df_also_adds look like self.df_excludes['col_exclude_also]
            df_also_adds.rename(columns={'Keyword': 'col_exclude_also'}, inplace=True)
            df_also_adds.drop(columns=['Count'], axis=1, inplace=True)

            # Reduce df_exclude_also to only contain the 'col_exclude_also' column
            df_exclude_also = self.df_excludes.copy(deep=True)
            df_exclude_also.drop(columns=['col_exclude_always'], axis=1, inplace=True)
            # Drop the NaN values that padded the column so it was the same length as
            # the 'col_exclude_always' column.
            df_exclude_also.dropna(axis=0, how='all', inplace=True)

            # Add new exclude words to exclude_also column.
            df_exclude_also = pd.concat([df_exclude_also, df_also_adds], ignore_index=True)
            
            # Drop the current column in self.df_excludes
            self.df_excludes.drop(columns=['col_exclude_also'], axis=1, inplace=True)
            # The next three lines will add NaNs to the columns to ensure they are the
            # same length.  If they are not they will not join properly or at all.
            max_len = max(len(self.df_excludes), len(df_exclude_also))
            self.df_excludes = self.df_excludes.reindex(range(max_len))
            df_exclude_also = df_exclude_also.reindex(range(max_len))
            # Add the alsos back in as a 'new' column.
            self.df_excludes['col_exclude_also'] = df_exclude_also['col_exclude_also']
          
            if num_rec_dropped == 1:
                print(f'{num_rec_dropped} record removed from the Counts sheet and the',
                     'Keyword was added to the also column.')
            else:
                print(f'{num_rec_dropped} records removed from the Counts sheet and their',
                     'Keywords were added to the also column.')
            return num_rec_dropped    
        else:
            print("\nNo new words found to add to the 'col_exclude_also' column.\n")
            return 0


    def process_words(self):
        '''After moving new excluded words, remove excluded words'''
        # Start by moving any new word to the also column
        words_moved = self.move_words_to_exclude_also()
        excludes_set = self.create_excludes_set()
        original_count = len(self.df_relevant)
 
        df_dropped = self.df_relevant[self.df_relevant['Keyword'].isin(excludes_set)].copy(deep=True)
        if len(df_dropped) > 0:
            # Delete rows from Relevant seet
            self.df_relevant = self.df_relevant[~self.df_relevant['Keyword'].isin(excludes_set)]
            # Format df_dropped to match Irrelevant sheet
            df_dropped.drop(columns=['Exclude'], axis=1, inplace=True)
            
            # If none were move but some matched the exclude list, empty the irrelevant dataframe.
            if words_moved == 0:
                self.df_irrelevant = df_dropped
            else:
                # # The next three lines will add NaNs to the columns to ensure they are the
                # # same length.  If they are not they will not join properly or at all.
                self.df_irrelevant = pd.concat([self.df_irrelevant, df_dropped], ignore_index=True)
        
        final_count = len(self.df_relevant)
        recs_dropped = original_count - final_count
        recs_dropped += words_moved
        
        if recs_dropped == 0:
            print('No records were deleted or moved.')
        else:
            with pd.ExcelWriter(self.excel_file) as writer:
                # Add the sheets back in
                self.df_original.to_excel(writer, sheet_name='Original', index=False)
                self.df_clean.to_excel(writer, sheet_name='Clean', index=False)
                self.df_words.to_excel(writer, sheet_name='Words', index=False)
                self.df_counts.to_excel(writer, sheet_name='Counts', index=False)
                self.df_irrelevant.to_excel(writer, sheet_name='Irrelevant', index=False)
                self.df_relevant.to_excel(writer, sheet_name='Relevant', index=False)
                #Add in the modified excludes
                self.df_excludes.to_excel(writer, sheet_name='Excludes', index=False)
            if recs_dropped == 1:
                print(f'{recs_dropped} record was deleted or moved.')
            else:    
                print(f'{recs_dropped} were deleted or moved.')

                


