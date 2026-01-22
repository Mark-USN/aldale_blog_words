
import numpy as np
import re
from openpyxl.descriptors import NoneSet
import torch
import logging
import time
from sentence_transformers import SentenceTransformer, util
model = SentenceTransformer("multi-qa-MiniLM-l6-cos-v1")

class Group_Transform_Exception(Exception):
    """A custom exception class."""
    pass

class Group_Transform(object):
    '''Compute the cosine for all the combinations of all sentences and build 
    the desired matrices and list used to create various spreadsheets.
    '''
    
    def __init__(self, records, consolidate_level, threshold=0.75, header=False):
        '''Create the object, save the records and possibly the record header.'''
        
        '''Also sets the initial threshold and creates various sentence lists
        from the record information.
        '''
        
        if header:
            self.rec_header = records.pop(0)
        self.records = records
        self.threshold = threshold
        self.cnsldtd_lvl = consolidate_level
        # Adjust row entries for the worksheet and header offset.
        # 1 because the worksheet starts at 1 and not 0
        # 1 because of the header row in the converted sheet.
        self.excel_row_offset = 2
        self.create_sentences_lists()
        
    def get_records_count(self):
        ''' The number of sentences we are working with. This squared will be 
        the size of the matrix to be computed. 
        '''
        
        return len(self.records)
        

    def compute_embeddings(self, clean=False):
        '''Computes the list of vectors for each sentence.'''
        
        # Encode all sentences
        if(clean):
            self.embeddings = model.encode(self.clean_sentences)
        else:
            self.embeddings = model.encode(self.sentences)
        
    def compute_cosines(self):
        '''Creates a simetrical square matrix comparing all the sentences to 
        eachother and returning a single cosine value for each comparison.
        '''
        
        if not hasattr(self, 'embeddings'):
            self.compute_embeddings()
        # Compute cosine similarity across all pairs
        cosines = util.cos_sim(self.embeddings, self.embeddings)
        self.cosines = cosines.detach().numpy() 
        # Just to clean up imprecisions from float operations.
        np.fill_diagonal(self.cosines, 1.0)
        # zero is typically used as a null value in some matrix operations (like 
        # np.tril()) so shift any 0's to negative three times the minimum value 
        # this machine will recognize.
        zero_locs = np.argwhere(self.cosines == 0.0)
        if zero_locs.size > 0:
            self.cosines[zero_locs] = -3 * np.finfo(float).eps
        # logging.debug(f'\nself.cosines:\n{self.cosines}\n')   
        # logging.debug(f'\nType of cosines = {type(self.cosines)}\n')
           

    def get_cosine_sheet(self):
        '''Returns the processed cosines matrix with sentences prepended to each
        row and the sentences that are used as headers in the spreadsheet.
        '''
        
        if not hasattr(self,'cosines'):
            self.compute_cosines()     
        cos_data = self.convert_rows_for_excel(self.cosines, header=False, titles=True)
        return (cos_data, self.sentence_titles)

        
    def create_lower_triangle_matrix(self):
        '''Remove the upper trianglar half of the cosines matrix and the 
        diagonal.
        '''
        
        '''These values are redundant as the same comparisons are in the lower
        trianglar half of the matrix and the diagonal compares each sentence 
        with itself, which should equal 1.
        '''
        
        if not hasattr(self, 'cosines'):
            self.compute_cosines()
        # # Leave out the diagonal all ones entry
        self.tri = np.tril(self.cosines, k=-1)
        # logging.debug(f'self.tri:\n{self.tri}\n')   


    def get_lower_triangle_sheet(self):
        '''Returns the processed lower triangle matrix with row titles and the
        sentence list for the spreadsheet header information.
        '''
        
        if not hasattr(self,'tri'):
            self.create_lower_triangle_mtrx()
        tri_data = self.convert_rows_for_excel(self.tri, header=False, titles=True)
        return (tri_data, self.sentence_titles)

        
    def sort_lower_triangle_matrix(self):
        '''Sort the cosines in decending order.'''
        
        '''This returns the sorted values and a matrix of lookup values that are
        used, given the row and column numbers, to find the row position in the 
        original matrix that this particular cosine value came from.  The column
        number is the same for both matrices. The column number and looked up 
        row number are also the indexes into the sentences list for the source
        and target sentences.
        '''

        if not hasattr(self, 'tri'):
            self.create_lower_triangle_mtrx()
        tmp_lu = np.argsort(self.tri, axis=0)[::-1, :]
        # Reorder the rows for each column of the original array based on the sorted indices
        tmp_mtrx = np.take_along_axis(self.tri, tmp_lu, axis=0)
        
        '''The decending sort results it values being in the order Positive, 
        Zero, Negative but numpy uses 0 as a null value when it does things like
        create the lower triangle matrix. So in compute_cosines() method we 
        replaced any zero values in the matrix with a small negative value to 
        avoid confusion and here we move the sorted values so the order is 
        Positive, Negative, Zeros. Finally, Sort messed up our null values in 
        the lookup matrix by replacing them with the row they were in.  Since 
        we know the last entry in column 0 should be ignored we zero it out 
        and in column 1 the bottom two rows should be 0 and so on to the last 
        column making the lookup table match what the ingnored rows should be 
        for each column.
        '''
        
        self.srtd, self.srtd_lu = self.reorder_sort(tmp_mtrx, tmp_lu)
        for col in np.arange(self.srtd_lu.shape[1]):
            rows = self.srtd_lu[:, col]
            locs = np.argwhere(rows <= col)
            if locs.size > 0:
                rows[locs] = 0
                
        # logging.debug(f'\nself.srtd:\n{self.srtd}\n')   
        # logging.debug(f'\nself.srtd_lu:\n{self.srtd_lu}\n')   

        
    def get_sorted_cosine_sheet(self):
        ''' Convert the matrix to the format excel expects and pass it and the 
        header values back.
        '''
        
        if not hasattr(self,'srtd'):
            self.sort_lower_triangle_mtrx()
        return (self.convert_rows_for_excel(self.srtd, header=False, titles=False), self.sentence_titles)

        
    def get_sorted_lookup_sheet(self):
        ''' Convert the lookup table to Excel format and pass it and the header
        information back.
        '''
        
        if not hasattr(self,'srtd_lu'):
            self.sort_lower_triangle_mtrx()
        sheet_sort_lu = self.srtd_lu + self.excel_row_offset
        mask = sheet_sort_lu == 2
        if np.any(mask):
            sheet_sort_lu[mask] = 0
        return (self.convert_rows_for_excel(sheet_sort_lu, header=False, titles=False), self.sentence_titles)

        
    def create_best_matrix(self):
        ''' Keep the highest value for each sentence and zero out all other 
        entries.
        '''
        
        if not hasattr(self, 'srtd'):
            self.sort_lower_triangle_mtrx()
            
        self.bst, self.bst_trck = self.keep_best_correlations(self.srtd, self.srtd_lu)

        # logging.debug(f'\nbst:\n{self.bst}')
        # logging.debug(f'\nbst_trck:\n{self.bst_lu}')


    def get_initial_topic_count(self):
        ''' The count of columns left after keeping only the highest cos values
        and removing any columns that have all zero entries.
        '''
        
        if not hasattr(self, 'bst'):
            self.create_best_matrix()
        return self.bst.shape[1]
        

    def get_best_cosine_sheet(self):
        ''' Format the best cosine matrix for Excel and return it and it's 
        header information.
        '''
        
        ''' Since columns have been removed from the matrix the header will 
        consist of two rows the first being the sentences and the second being
        the column they came from.
        '''
        
        if not hasattr(self,'bst'):
            self.create_best_matrix()
        # Add worksheet offsets.     
        sheet_bst_trck = self.bst_trck + self.excel_row_offset

        return (self.convert_rows_for_excel(self.bst, header=False, titles=False), 
                self.create_tracking_sheet_header(sheet_bst_trck))

        
    def get_best_tracking_sheet(self):
        ''' Format the array for Excel and return it and the header. '''
        
        if not hasattr(self,'bst_trck'):
            self.create_best_matrix()
        # Add worksheet offsets.     
        sheet_bst_trck = self.bst_trck + self.excel_row_offset
        bst_trck_data = sheet_bst_trck[1:, :]
        mask = bst_trck_data == 2
        if np.any(mask):
            bst_trck_data[mask] = 0
        return (self.convert_rows_for_excel(bst_trck_data, header=False, titles=False), 
                self.create_tracking_sheet_header(sheet_bst_trck))

        
    def create_consolidated_matrix(self):
        ''' Move the information in any columns containing self.cnsldtd_lvl 
        rows or less to the row containing the column index for this column.
        '''
        
        ''' After consolidation the Lower Triangluar matrix can no longer be 
        guarnteed to have a vaild value for the indexes passed in. A sentence
        can appear under its own column and a row number may be less then the
        column number given.  
        '''
        
        if not hasattr(self,'bst_trck'):
            self.create_best_matrix()
        self.cnsldtd_trck = self.bst_trck.copy(order='C')
        self.cnsldtd_trck = self.eleminate_short_rows(self.cnsldtd_trck)
        cnsldtd_data = self.cnsldtd_trck[1:, :]
        # should not be needed.
        self.cnsldtd_trck[1:, :] = self.move_values_up(cnsldtd_data)
        # Remove all the columns we just emptied from the tracking matrix.
        self.cnsldtd_trck = self.remove_tracking_columns(self.cnsldtd_trck)
        self.cnsldtd_trck = self.sort_tracking_columns_by_cosine(self.cnsldtd_trck)
        self.cnsldtd = self.create_cosines_from_tracking_index(self.cnsldtd_trck)
        # remove rows with all zeros.
        self.cnsldtd, self.cnsldtd_trck = self.truncate_tracking_and_matrix_rows(self.cnsldtd, self.cnsldtd_trck)


    def get_consolidated_topic_count(self):
        ''' Return the number of columns left in the array after consolidation.'''
        
        if not hasattr(self, 'cnsldtd'):
            self.create_consolidated_matrix()
        return self.cnsldtd.shape[1]


    def get_consolidated_cosine_sheet(self):
        ''' Format consolidated cosine array for Excel and return it and it's
        header information.
        '''
        
        if not hasattr(self,'cnsldtd'):
            self.create_consolidated_matrix()
            
        # Add worksheet offsets for converted header row.     
        sheet_cnsldtd_trck = self.cnsldtd_trck + self.excel_row_offset
            
        return (self.convert_rows_for_excel(self.cnsldtd, header=False, titles=False), 
                self.create_tracking_sheet_header(sheet_cnsldtd_trck))
    

    def get_consolidated_tracking_sheet(self):
        ''' Format the tracking data for Excel and return it and it's header.'''
        
        if not hasattr(self,'cnsldtd_trck'):
            self.create_consolidated_matrix()
        # Add worksheet offsets.     
        sheet_cnsldtd_trck = self.cnsldtd_trck + self.excel_row_offset
        cnsldtd_trck_data = sheet_cnsldtd_trck[1:, :]
        mask = cnsldtd_trck_data == 2
        if np.any(mask):
            cnsldtd_trck_data[mask] = 0
        return (self.convert_rows_for_excel(cnsldtd_trck_data, header=False, titles=False), 
                self.create_tracking_sheet_header(sheet_cnsldtd_trck))
    

    def get_consolidated_sentences_sheet(self):
        ''' Create a list of the Topic (column) sentences and the sentences 
        under them.
        '''
        
        if not hasattr(self, 'cnsldtd_trck'):
            self.create_consolidated_matrix()
        # Add worksheet offsets.     
        sheet_cnsldtd_trck = self.cnsldtd_trck + self.excel_row_offset
        return (self.create_tracking_sheet_sentences(self.cnsldtd_trck),
                    self.create_tracking_sheet_header(sheet_cnsldtd_trck))


    def create_topic_matrix(self):
        ''' Gets the sentences referred to in the first row of the tracking 
        matrix and converts them to a column.
        '''
        
        if not hasattr(self, 'cnsldtd_trck'):
            self.create_consolidated_matrix()
            
        # Extract the first row from the tracking matrix
        topics = np.array([self.sentences[i] for i in self.cnsldtd_trck[0, :]])

        topics_col = topics.reshape(-1, 1)
        self.topics = topics_col



    
    def get_topic_sheet(self):
        ''' Format the column of sentences for Excel and return them and the 
        header.
        '''
        
        if not hasattr(self, 'topics'):
            self.create_topic_matrix()
        return (self.convert_rows_for_excel(self.topics, header=False, titles=False), ['Topics'])

    
    def create_aldos_list(self, sheet=False):
        ''' Create the information in the format Aldo requested. Since the rows
        consist of a variety data types, arrays can not be used.
        '''
        
        if not hasattr(self, 'cnsldtd_trck'):
            self.create_consolidated_matrix()
            
        from_blw = False
        self.aldos_lst = []
        # Create the header
        self.aldos_header = [['Source', 'Topic', 'Topic_Volume', 'Source_Volume', 'Cosine', 'Source_row', 'Sort_Order']]
        num_recs = len(self.records)
        
        if sheet:
            src_row = self.excel_row_offset
        else:
            src_row = 0
            
        # 0 is the 'null' value in the matrices and the comparison of the first
        # sentence with itself is not in the matrix so handle it here.
        self.aldos_lst.append([self.records[0][0], self.records[0][0],
                            self.records[0][2],self.records[0][2],
                            1.0, src_row, False])

        for i in range(0, num_recs):
            src_sentence = self.records[i][0]
            src_vol = self.records[i][2]
            topic = ''
            topic_vol = -2
            topic_cos = -2
            # The first record was added above so add one to the index.
            if sheet:
                src_row = i + self.excel_row_offset + 1
            else:
                src_row = i + 1
            # Find the column index where the value is found skipping the
            # column index row of the tracking matrix
            topic_cols = self.cnsldtd_trck[0, :]
            trck_data = self.cnsldtd_trck[1:, :]
            src_loc = np.argwhere(trck_data == i)
            if src_loc.shape[0] > 0: 
                row, col = src_loc[0]
                topic_col = topic_cols[col]
                topic = self.records[topic_col][0]
                topic_vol = self.records[topic_col][2]
                topic_cos = self.cosines[i, topic_col]                        
                from_blw = topic_cos < self.threshold
            self.aldos_lst.append([src_sentence, topic, topic_vol, src_vol, topic_cos, src_row, from_blw])
            
        # Sort by topic volume then by topic, whether above or below cutoff, 
        # source volume, and finally cosine all from highest to lowest value.
        # TODO: Seems to put record 0 after its below cutoff values in the 
        # sheet. Other 'Above Cutoff' values appear normally.     
        self.aldos_lst = sorted(self.aldos_lst, key=lambda x: (x[2], x[1], ~x[6], x[3], x[4]), reverse=True)
        # Replace from_blw with the sort order
        for i, rec in enumerate(self.aldos_lst):
            rec[6] = i
                            

    def get_aldos_sheet(self):
        ''' The list is already in the format Excel expects so return it and the
        header.
        '''
        
        if not hasattr(self, 'aldos_lst'):
            self.create_aldos_list(sheet=True)
        return self.aldos_lst, self.aldos_header

    '''Testing Methods. '''            

    if __debug__:
        def check_all_lookup_tables(self):
            ''' Driver method to run checks on the tracking and lookup matrices
            that are created in the program. 
            '''
            
            if not hasattr(self, 'cnsldtd_trck'):
                self.create_consolidated_matrix()
            self.check_lookup_table_integrity(self.tri, self.srtd, self.srtd_lu, 'srtd_lu')
            self.tracking_matrix_counts_and_content(self.bst_trck, 'bst_trck')
            self.check_tracking_table_integrity(self.tri, self.bst, self.bst_trck, 'bst_trck')
            self.tracking_matrix_counts_and_content(self.cnsldtd_trck, 'cnsldtd_trck')
            self.check_tracking_table_integrity(self.cosines, self.cnsldtd, self.cnsldtd_trck, 'cnsldtd_trck') 
            logging.debug('check_all_lookup_tables completed successfully')
 
            
        def check_lookup_table_integrity(self, src_mtrx, tgt_mtrx, lu_mtrx, title):
            ''' Check the cosine value in the tgt_mtrx against the value in the 
            src_mtrx pointed to by the value of pointed to by the row and column
            entries and the column number in the lookup array.
            '''
            
            for col in np.arange(lu_mtrx.shape[1]):
                for row in np.arange(0, lu_mtrx.shape[0] - col - 1):  # Only iterate up to the diagonal
                    org_row = lu_mtrx[row, col]
                    target_cosine = tgt_mtrx[row, col]
                    assert (
                        src_mtrx[org_row, col] == target_cosine
                    ), (
                        f'Lookup Table {title} failed position ({org_row}, {col}) != ({row}, {col})'
                    )
            logging.debug(f'check_lookup_table for {title} completed successfully.')


        def check_tracking_table_integrity(self, src_mtrx, tgt_mtrx, lu_trck_mtrx, title):
            ''' Basically the same as the lookup version except row 0 contains 
            the original column number.
            '''
            
            for col in np.arange(lu_trck_mtrx.shape[1]):
                for row in np.arange(1, lu_trck_mtrx.shape[0] - col - 1):  # Only iterate up to the diagonal
                    org_col = lu_trck_mtrx[0, col]
                    org_row = lu_trck_mtrx[row, col]
                    target_cosine = tgt_mtrx[row - 1, col]
                    assert (
                        src_mtrx[org_row, org_col] == target_cosine
                    ), (
                        f'Tracking Table failed position ({org_row}, {org_col}) != ({row}, {org_col})'
                    )
            logging.debug(f'check_tracking_table for {title} completed successfully.')


        def tracking_matrix_counts_and_content(self, tracking, title):
            ''' Make sure the tracking matrix has the right number of entries
            and that none are missing.
            '''
            trck_data = tracking[1:, :]
            mask = trck_data != 0
            index_lst = trck_data[mask].flatten().tolist()
                
            unique_indexs = set(index_lst)
            # Get the number of rows in the matrices (assuming both matrices have the same shape)
            rec_len = len(self.records)
            # Define the expected range from 1 to rec_len-1
            # Comparison of sentence 0 with itself is not in the table so don't
            # worry about it.
            expected_values = set(range(1, rec_len))    
            # Find any missing values
            missing_values = expected_values - unique_indexs
    
            # Check for duplicates and missing values
            assert (
                len(index_lst) == len(unique_indexs)
            ), (
                f'{title} Matrix length {len(index_lst)} does not match unique Values set size {len(unique_indexs)}.'
            )
            assert not missing_values, f'{title} Values are missing from the Combined Matrix. The difference is {missing_values}.'


    '''Helper Methods. '''            


    def move_values_up(self, matrix):
        ''' For each column move non-zero values to the top of the list and fill
        with zeros.
        '''
        
        ''' When used for tracking matrices row 0 is not passed in. Also after 
        the move if the first entry is zero all the entries below it are zero
        and the column is empty and can be removed.
        '''
        
        rows, cols = matrix.shape    
        for col in np.arange(cols):
            # Get the column
            column = matrix[:, col]    
            # Find the non-zero elements in the column
            non_zero_elements = column[column != 0]
            # Count the number of non-zero elements
            num_non_zeros = non_zero_elements.size
            # Calculate the number of zeros to append
            num_zeros = rows - num_non_zeros
            # Construct the new column
            new_column = np.concatenate((non_zero_elements, np.zeros(num_zeros)))
            # Assign the new column back to the matrix
            matrix[:, col] = new_column
        return matrix
            

    # Function to zero out all occurrences except the one with the highest
    # cosine correlation.
    def keep_best_correlations(self, data_mtrx, data_lu):
        ''' Return array containing only the highest cosine value found for 
        each sentence and its tracking matrix.
        '''

        ''' Here we get the location for each array entry containing the 
        sentence index for each sentence. There will be an entry in every
        column where the column index is less than the sentence index.
        the results are then searched to find the location of the maximum value
        for a sentence in the table. This location and value are saved and all
        the locations found are set to zero (the null value) then the saved 
        value is written back to its original location.  
        Finally all the valid data in each column is place above any null (zero)
        values and the array is 'compacted' removing any rows that are all 
        zeros, and any columns that are zero below the column tracking row (row 
        0 in the tracking matrix).  The resultant matrix contains only the sent-
        ence cosines with the highest value across the original array.  
        Note that since the diagonal was not included there is no entry for 
        sentence 0 since it would only be compared with itself and should ewual
        1.0. Also the last column will have no entries since it has already been
        comparied with every sentence below it in the array.
        '''

        rtn_mtrx = data_mtrx.copy(order='C')
        lu = data_lu.copy(order='C')
        num_recs = len(self.records)    # lu.shape[0]
        # 1. Zero is used as a null value in the lookup table and does not appear 
        # as a valid value as the lower triangle was taken below the diagonal.
        # 2. One has only one instance and appears only in column 0 of the lookup
        # table so we dont need to evaluate it either. 
        # 3. The number of rows or columns of the lookup table is the same as 
        # the number of sentences that were processed. 
        for tgt_ndx in range(2, num_recs):
            # Get a boolean array of where the sentence_index appears in the current and subsequent columns
            mask = (lu == tgt_ndx)

            if np.any(mask):
                # Get the indices of these matches
                row_indices, col_indices = np.nonzero(mask)
                # Find the maximum cosine value among the matches
                matching_values = rtn_mtrx[row_indices, col_indices]
                max_value_index = np.argmax(matching_values)
                max_row = row_indices[max_value_index]
                max_col = col_indices[max_value_index]
                max_value = matching_values[max_value_index]
                    
                # Zero out everything then write the value back 
                # where it belongs
                lu[row_indices, col_indices] = 0
                rtn_mtrx[row_indices, col_indices] = 0.0
                # Keep the highest value in the other matrix
                lu[max_row, max_col] = tgt_ndx
                rtn_mtrx[max_row, max_col] = max_value

        rtn_mtrx = self.move_values_up(rtn_mtrx)
        lu = self.move_values_up(lu)
        rtn_mtrx, lu = self.truncate_lookup_rows(rtn_mtrx, lu)
        if rtn_mtrx.shape[0] != 0:
            rtn_trck = self.create_tracking_table(lu)
            rtn_trck = self.remove_tracking_columns(rtn_trck)
            rtn_mtrx = self.remove_matrix_columns(rtn_mtrx, rtn_trck)
        else:
            rtn_mtrx = None
            rtn_trck = None
        
        return (rtn_mtrx, rtn_trck)
        

    def sort_tracking_columns_by_cosine(self, tracking_matrix):
        ''' Sort each column in the tracking array based on its values in the
        associated cosine array and return the resultant tracking array.
        '''
        
        max_recs = len(self.records)
        
        # Get the number of columns (excluding the index row)
        # num_cols = tracking_matrix.shape[1]
        org_cols = tracking_matrix[0, :]
        # Iterate over each column and sort based on the associated volume and cosine values
        # for col in range(num_cols):
        for col in np.arange(tracking_matrix.shape[1]):
            # Get the data column (skip the first row which is the index row)
            column_data = tracking_matrix[1:, col]
            # cosines_array = np.array(self.cosines[:, org_cols[col]])
            cosines_array = self.cosines[:, org_cols[col]]

            # Create an array to hold the sorting keys
            cosine_keys = np.array([cosines_array[val] if val != 0 else -float('inf') for val in column_data])
        
            sorted_indices = np.argsort(cosine_keys, axis=0)[::-1]
            tracking_matrix[1:, col] = column_data[sorted_indices]
    
        return tracking_matrix


    def eleminate_short_rows(self, tracking):
        ''' For each row whose length is less than or equal to self.cnsldtd_lvl,
        move the columns rows to the column containing the tracking column entry
        in its data and zero out the original column.
        '''
                    
        rtn_trck = tracking
        org_cols = rtn_trck[0, :]
        data = rtn_trck[1:, :]
        lens = np.count_nonzero(data, axis=0)

        # Make sure we have some 'room' to work with in the columns.
        if self.cnsldtd_lvl <= 0:
            return rtn_trck
        elif self.cnsldtd_lvl+2 >= lens.max():
            return rtn_trck
        
        # The columns we want to delete
        src_cols = np.argwhere(lens <= self.cnsldtd_lvl)    
        if src_cols.size > 0: 
            # Moving across the columns from right to left will avoid future 
            # changes from affecting rows we have already cleared.
            for col in src_cols[:, 0][::-1]:
                # Don't rely on cnsldtd_lens info as previous interations could
                # have added values to these rows.
                src_rows_used = np.count_nonzero(data[:, col], axis=0)
                # If the column's data has grown this much, lets keep it.
                if src_rows_used/3.0 >= self.cnsldtd_lvl:
                    continue
                dest_loc = np.argwhere(data == org_cols[col])
                if dest_loc.size > 0:
                    dest_col = dest_loc[0][1]
                    data_rows = data.shape[0]
                    dest_rows_used = np.count_nonzero(data[:, dest_col], axis=0)
                    empty_rows = data_rows - dest_rows_used
                    needed_rows = data_rows - empty_rows + dest_rows_used
                
                    if empty_rows < src_rows_used:
                        rtn_trck.resize((needed_rows + 5 * self.cnsldtd_lvl,
                                         rtn_trck.shape[1]), refcheck=False)
                        org_cols = rtn_trck[0, :]
                        data = rtn_trck[1:, :]  

                    if src_rows_used > 1:
                        src_values = data[0:src_rows_used, col]
                        first_row = dest_rows_used
                        last_row = dest_rows_used + src_rows_used
                        data[first_row:last_row, dest_col] = src_values
                        data[0:src_rows_used, col] = 0
                    else:
                        src_value = data[0, col]
                        data[dest_rows_used, dest_col] = src_value
                        data[0, col] = 0                            
        return rtn_trck


    def create_cosines_from_tracking_index(self, trck):
        ''' Use the given tracking array to populate a cosine array.'''
        
        # The information has been moved in the tracking matrix, now use
        # matrix to build the cosines matrix.                
        # Determine the shape of the resultant matrix
        num_rows = trck.shape[0] - 1
        num_cols = trck.shape[1]
        # Initialize the result matrix with zeros
        cosines = np.zeros((num_rows, num_cols), dtype=self.tri.dtype)
        # Extract the columns to use from the original matrix
        original_columns = trck[0, :]
        # Iterate over each column in the tracking matrix
        for col in range(num_cols):
            # Get the original column index
            orig_col = original_columns[col]
            # Get the original row locations for this column
            orig_rows = trck[1:, col]
            # Place the values in the correct positions in the result matrix
            # Use self.cosines as consolidated array could break expectations 
            # of the Lower Triangle Matrix especially as it does not include 
            # the diagonal.
            cosines[:, col] = self.cosines[orig_rows, orig_col].copy(order='C')
        return cosines

    def reorder_sort(self, matrix, index_array):
        ''' Change the original decending sort from positive values, 
        zero values, negative values to positive, negative, zeros.
        '''
        
        ''' This is necessary since zero represents a null value in the index.'''
        
        nrows, ncols = matrix.shape
        reordered_mtrx = np.zeros_like(matrix)
        reordered_index_array = np.zeros_like(index_array)
    
        for col in range(ncols):
            column = matrix[:, col]
        
            # Separate the positive, negative, and zero values and get their indices
            negatives = np.where(column < 0)[0]
            positives = np.where(column > 0)[0]
            zeros = np.where(column == 0)[0]
        
            # Combine the indices in the order: negatives, positives, zeros
            new_order = np.concatenate((positives, negatives, zeros))
        
            # Apply the new order to the matrix and the index array
            reordered_mtrx[:, col] = column[new_order]
            reordered_index_array[:, col] = index_array[new_order, col]
    
        return reordered_mtrx, reordered_index_array
 

    def remove_tracking_columns(self, tracking):
        ''' Remove any columns where the row 1 value for that column is zero.'''
        
        ''' If the tracking array is properly sorted all rows below the first 
        zero row of a column will also be zero.  So, since row 0 is the column
        index, if row 1 is zero the column is empty.
        '''
        
        rtn_trck = tracking.copy(order='C')
        columns_to_delete = [col for col in range(rtn_trck.shape[1]) 
                             if rtn_trck[1, col] == 0]
        # Create a mask to keep only the rows that are not to be deleted
        mask = np.ones(rtn_trck.shape[1], dtype=bool)
        mask[columns_to_delete] = False
        # Apply the mask to the matrix to filter out the rows to be deleted
        rtn_trck = rtn_trck[:, mask]
        return rtn_trck


    # Assumes columns have already been removed from the tracking_index
    def remove_matrix_columns(self, matrix, tracking_index):
        ''' Remove columns in the cosines array that are not in the tracking 
        array.
        '''
        
        ''' Oboviously assumes the columns have aready been removed from the 
        tracking array.
        '''
        
        # first_row = tracking_index[0].astype(float)
        rtn_mtrx = np.zeros((tracking_index.shape[0]-1, tracking_index.shape[1]))
        # rtn_mtrx[0] = first_row
        # Get the column indices from the first row of tracking_index
        col_indices = tracking_index[0]
        row_indices = tracking_index[1:, :]
        # Use broadcasting and advanced indexing to fill values matrix
        rtn_mtrx = self.tri[row_indices, col_indices]
        return rtn_mtrx


    def create_tracking_sheet_sentences(self, tracking_index):
        ''' Use the tracking array data as indexes into the sentence array.'''
        
        # Use numpy's vectorized indexing to create the new matrix
        # First, convert sentences list to a numpy array
        sentence_array = np.array(self.sentences)
        mtrx_data = tracking_index[1:, :]
        # Indexing to create the new matrix
        rtn_mtrx = sentence_array[mtrx_data]
        
        rows, cols = mtrx_data.shape    
        for row in range(rows):
            row_data = mtrx_data[row, :]
            # Find the column indices where the value is zero
            # zero_indices = np.where(row_data == 0)[0]
            zero_indices = np.where(row_data == 0)
            # Set the corresponding columns to an empty string for this row in the other
            # matrix
            if len(zero_indices) > 0:
                rtn_mtrx[row, zero_indices] = ''
        return self.convert_rows_for_excel(rtn_mtrx, header=False, titles=False)
            

    def create_tracking_sheet_header(self, tracking_index):
        ''' Return a row of column sentences and a row of their indexes.'''
        
        # Use numpy's vectorized indexing to create the new matrix
        # First, convert sentences list to a numpy array
        sheet_indexes = tracking_index[0, :].tolist()
        real_indexes = [val - self.excel_row_offset for val in sheet_indexes]
        sentence_array = np.array(self.sentence_titles)
        # Indexing to create the new matrix
        sentences = sentence_array[real_indexes].tolist()
        rtn_header = []
        rtn_header.append(sentences)
        rtn_header.append(sheet_indexes)
        
        return self.convert_rows_for_excel(rtn_header, header=False, titles=False)
            

    def truncate_lookup_rows(self, matrix, index):
        ''' Truncates the matrix to remove all rows containing only zeros.'''
        
        '''Assumes the first row containing all zeros and all rows below it 
        are all zeros.
        '''
        
        zeros_mask = np.all(index == 0, axis=1)
        if np.any(zeros_mask):
            first_zero_row_index = np.argmax(zeros_mask)
            if first_zero_row_index > 0:
                # Resize the array in place
                matrix.resize((first_zero_row_index, matrix.shape[1]), refcheck=False)
                index.resize((first_zero_row_index, index.shape[1]), refcheck=False)
            else:
                matrix = None
                index = None
        return matrix, index


    def truncate_tracking_and_matrix_rows(self, matrix, tracking):
        ''' Truncates the matrix to remove all rows containing only zeros.'''
        
        ''' Assumes the first row containing all zeros and all rows below it 
        are all zeros.
        '''
        
        # Identify rows that are not all zeros
        non_zero_rows = tracking[~np.all(tracking == 0, axis=1)]
        # Modify the original array in place by resizing it
        tracking = non_zero_rows          # .resize(non_zero_rows.shape, refcheck=False)
        matrix = self.create_cosines_from_tracking_index(tracking)
        
        return matrix, tracking


    ''' Any change to the column structure requires a 'tracking' matrix in 
    order to retain the original source sentence of the comparison the cosines
    represent. '''
    
    def create_tracking_table(self, lu_table):
        ''' Prepend a row of sentence indexs representing the columns.'''
        
        # Create the row containing the column indices
        column_indices = np.arange(lu_table.shape[1])
        # Stack the column indices row on top of the original array
        rtn_table = np.vstack((column_indices, lu_table))
        return rtn_table
       

    def convert_rows_for_excel(self, array_, header=True, titles=True):
        ''' Convert array to list adding header and row titles if needed.'''
        
        rows = []
        if titles:
            row_titles = self.get_row_titles()
            rows = [row_titles[row_ndx] + row.tolist() for row_ndx, row in enumerate(array_)]
            if header:
                header_row = ['']
                header_row.extend(self.get_header())
                rows.insert(0, header_row)
        elif header:
            rows = [row.tolist() for row in array_]
            rows.insert(0, self.get_header())
        else:
            if type(array_) == list:
                rows = array_
            else:
                rows = [row.tolist() for row in array_]
           
        return rows
    

    def set_threshold(self, threshold):
        ''' Set the threshold to a new value, returning its original value.'''
        
        ''' Does not have much effect here as it can be changed in the workbook
        and will take effect there.
        '''
        
        if threshold > 1.0 or threshold < -1.0:
            print(f'Threshold must be between 1 and -1. You entered: {threshold}')
        else:
            old_threshold = self.threshold
            self.threshold = threshold
            return old_threshold


    def get_threshold(self):
        ''' Return the current threshold setting.'''
        
        return self.threshold


    def set_consolidation_level(self, level):
        ''' Set a new consolidation_level and return the current one.'''
        
        ''' The self.create_consolidated_matrix(), and subsequent building and 
        sheet creation methods need to be called to see the effects of this 
        change.
        '''
        
        old_level = self.cnsldtd_lvl
        self.cnsldtd_lvl = level
        return old_level


    def get_consolidation_level(self):
        ''' Return the current consolidation level. '''
        return self.cnsldtd_lvl


    def create_sentences_lists(self):
        '''Make a list of sentences, cleaned sentences and title sentences from 
        the records.
        '''
        
        self.sentences = [rec[0] for rec in self.records]
        self.clean_sentences = [re.sub(r'[^\w\s]', '', sentence.strip().casefold()) for sentence in self.sentences]
        self.sentence_titles = [re.sub(r'\s+','_', sentence) for sentence in self.clean_sentences]    
  

    def get_header(self):
        ''' Return a row of sentences to use as a worksheet header.'''
        
        return self.sentence_titles


    def get_row_titles(self):
        ''' Return a column of sentences to prepend to a row's data as the title
        for that row.
        '''
        
        sentences = []
        sentences = [[sentence] for sentence in self.sentence_titles]
        return sentences
