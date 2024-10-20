import pandas as pd
import tkinter as tk
from tkinter import ttk

# Read the Excel file, two paths are defined, Ellen's and Shawn's/Erick's
#file_path = 'C:/Users/ellen/Downloads/python/db_PARTING.xlsx'
file_path = '/Users/erickvillanueva/Documents/GitHub/Parting/db_PARTING.xlsx'  
df = pd.read_excel(file_path)
    
# Define label groups with corresponding columns from the DataFrame
label_groups = {
    'SPEC_FEAT': ['SPEC_FEAT'],
    'PRO_Nominal': ['PRO_Form', 'Basis', 'withP_complement', 'withQ_complement', 'withN_complement'],
    'REL_Type': ['RELATION'],
    'OUT_SYN': ['DET', 'Q_COMPLEXITY', 'N_OUT', 'Q_INFL'],
    'OUT_SEM': ['SORT_TYPE', 'SEM_TARGET'],
    'INN_Syn': ['Number', 'Case', 'Det', 'Det_Infl', 'Preposition', 'P_Sem'],
    'ADJ': ['Adj', 'Adj_Infl', 'Adj_Comparison'],
    'INN_Lex': ['ONTO_type', 'PERCEP_type', 'GRANU_type', 'NUM_type'],
    'INN_ctxt': ['ctxt_ENTITY', 'ctxt_PERCEP', 'exclamative']
}

class PhraseMatching:
    def __init__(self, root):
        self.root = root
        self.root.title("PARTING")

        # Main frame setup
        main_frame = tk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=1)

        # Canvas setup for scrolling
        canvas = tk.Canvas(main_frame)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=1)

        # Vertical and horizontal scrollbar
        scrollbar_y = ttk.Scrollbar(main_frame, orient=tk.VERTICAL, command=canvas.yview)
        scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)
        scrollbar_x = ttk.Scrollbar(self.root, orient=tk.HORIZONTAL, command=canvas.xview)
        scrollbar_x.pack(side=tk.BOTTOM, fill=tk.X)

        # Configure scroll commands
        canvas.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)
        canvas.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

        # Frame to hold the content
        content_frame = tk.Frame(canvas)
        canvas.create_window((0, 0), window=content_frame, anchor="nw")

        # Dictionary to hold combo box selections
        self.selections = {}

        current_col = 0
        current_row = 0
        row_offset = 2
        max_columns = 6  # Set the maximum number of columns before wrapping to the next row

        # Create labels and combo boxes based on label groups
        for group_name, labels in label_groups.items():
            ttk.Label(content_frame, text=group_name, font=("Arial", 16, "bold")).grid(
                row=current_row, column=current_col, columnspan=2, padx=5, pady=5, sticky="n")
            
            for idx, label in enumerate(labels):
                # Label for each selection option
                ttk.Label(content_frame, text=label, font=("Arial", 13)).grid(
                    row=row_offset + current_row + idx, column=current_col, padx=5, pady=2, sticky="e")
                
                # Get unique values from DataFrame for each label
                values = df[label].dropna().unique().tolist()
                values = [str(v) for v in values]
                
                if df[label].isna().any():
                    values.append('None')

                # Create combo box for selections
                combo = ttk.Combobox(content_frame, values=['ALL'] + values, font=('Courier', 12), width=20)
                combo.grid(row=row_offset + current_row + idx, column=current_col + 1, padx=5, pady=2, sticky="w")
                combo.set('ALL')
                self.selections[label] = combo
            
            current_col += 2

            # Move to the next row if the max column limit is reached
            if current_col >= max_columns * 2:
                current_col = 0
                current_row += len(labels) + row_offset

        # Submit button to filter data
        max_group_height = len(max(label_groups.values(), key=len)) + row_offset
        submit_btn = ttk.Button(content_frame, text="Submit", command=self.filter_data)
        submit_btn.grid(row=current_row + max_group_height + 1, column=5, columnspan=2, pady=25)

        # Reset button to clear selections
        reset_btn = ttk.Button(content_frame, text="Reset", command=self.reset_filters)
        reset_btn.grid(row=current_row + max_group_height + 1, column=7, columnspan=2, pady=25)

        # Text widget to display final results and Label for total count
        # Set the Text widget to use a monospaced font for proper alignment
        self.final_result_text = tk.Text(content_frame, height=20, width=150, font=("Courier", 12))
        self.final_result_text.grid(row=current_row + max_group_height + 2, column=3, columnspan=current_col + 2, pady=25)
        self.final_result_count = ttk.Label(content_frame, text="Total Count: 0", font=("Arial", 20))
        self.final_result_count.grid(row=current_row + max_group_height + 3, column=3, columnspan=current_col + 2, pady=2)

        # Bind combo box selection events to apply disregard rules
        for group_name, labels in label_groups.items():
            for label in labels:
                self.selections[label].bind("<<ComboboxSelected>>", lambda e: self.apply_disregard_rules())

    def apply_disregard_rules(self):
        """ Apply rules to disregard certain selections based on user choices. """
        # Reset all selections to normal before applying the rules
        for group_name, labels in label_groups.items():
            for label in labels:
                self.selections[label].config(state='normal')
      
        # Rule 1: Disable all groups to the right of PRO-Nominal if a selection is made in PRO-Nominal
        pro_nominal_selected = any(self.selections[label].get() != 'ALL' for label in label_groups['PRO_Nominal'])
        if pro_nominal_selected:
            disable_groups = ['REL_Type', 'OUT_SYN', 'OUT_SEM', 'INN_Syn', 'ADJ', 'INN_Lex', 'INN_ctxt']
            for group in disable_groups:
                for label in label_groups[group]:
                    self.selections[label].config(state='disabled')
        
        # Rule 2: Disable Q-COMPLEXITY and Q-INFL if DET in OUT-SYN is set to 'None'
        if self.selections['DET'].get() == 'None':
            self.selections['Q_COMPLEXITY'].set('ALL')
            self.selections['Q_COMPLEXITY'].config(state='disabled')
            self.selections['Q_INFL'].set('ALL')
            self.selections['Q_INFL'].config(state='disabled')

        # Rule 3: Disable Det-Infl in OUT-SYN if Det is not 'Def' or 'IA'
        if self.selections['Det'].get() != 'Def' and self.selections['Det'].get() != 'IA':
            self.selections['Det_Infl'].set('ALL')
            self.selections['Det_Infl'].config(state='disabled')

        # Rule 4: Disable Preposition and P_Sem in INN-SYN if Case is not 'PP'
        if self.selections['Case'].get() != 'PP':
            self.selections['Preposition'].set('ALL')
            self.selections['Preposition'].config(state='disabled')
            self.selections['P_Sem'].set('ALL')
            self.selections['P_Sem'].config(state='disabled')
        else:
            self.selections['Preposition'].config(state='disabled')
            self.selections['P_Sem'].config(state='disabled')

        # Rule 5: Disable Adj-Infl and Adj-Comparison in ADJ if Adj is set to 'None'
        if self.selections['Adj'].get() == 'None':
            self.selections['Adj_Infl'].set('ALL')
            self.selections['Adj_Infl'].config(state='disabled')
            self.selections['Adj_Comparison'].set('ALL')
            self.selections['Adj_Comparison'].config(state='disabled')

    def filter_data(self):
        """ Filter the data based on user selections and display the results. """
        final_filtered_df = df.copy()

        # Apply disregard rules before filtering
        self.apply_disregard_rules()

        # Iterate through label groups and filter DataFrame based on selections
        for group_name, labels in label_groups.items():
            group_query_parts = []
            for label in labels:
                selection = self.selections[label].get()
                if selection != 'ALL':
                    if selection == 'None':
                        group_query_parts.append(f"{label}.isna()")
                    else:
                        group_query_parts.append(f"{label} == '{selection}'")

            if group_query_parts:
                group_query = " & ".join(group_query_parts)
                group_filtered_df = final_filtered_df.query(group_query)
            else:
                group_filtered_df = final_filtered_df
            
            # Merge the filtered DataFrame
            final_filtered_df = pd.merge(final_filtered_df, group_filtered_df, how='inner').drop_duplicates()

        # Prepare headers
        headers = ['INDEX', 'Meaning', 'Context', 'NP']
        
        # Format the results using format_output
        formatted_results = format_output(final_filtered_df[headers])

        # Update the result box
        self.final_result_text.delete(1.0, tk.END)
        self.final_result_text.insert(tk.END, formatted_results)
        
        final_count = len(final_filtered_df)
        self.final_result_count.config(text=f"Total Count: {final_count}")

        # Write the output to a text file
        #with open('C:/Users/ellen/Downloads/python/search.txt', 'w') as f:  # THIS HAS TO BE CHANGED TO THE DESIRED PATH 
        with open('/Users/erickvillanueva/Documents/GitHub/Parting/search.txt', 'w') as f:  # THIS HAS TO BE CHANGED TO THE DESIRED PATH 
            selected_options = {label: self.selections[label].get() for group_name, labels in label_groups.items() for label in labels}
            f.write("Selected Options:\n")
            f.write(str(selected_options) + "\n\n")
            f.write("Resulting Rows:\n\n")
            f.write(formatted_results)


    def reset_filters(self):
        """ Reset all selections to default values. """
        for group_name, labels in label_groups.items():
            for label in labels:
                self.selections[label].set('ALL')
                self.selections[label].config(state='normal')
        
        # Clear results display
        self.final_result_text.delete(1.0, tk.END)
        self.final_result_count.config(text="Total Count: 0")
        self.apply_disregard_rules()

def format_output(df):
    """ Set boundaries for column for better visibility in the results box and txt file. """
    # Set boundaries and header
    formatted_output = "{:<10} {:<40} {:<40} {:<30}".format('INDEX', 'Meaning', 'Context', 'NP') + "\n"
    for index, row in df.iterrows():
        # Shorten fields to 32 characters and append '[...]' if they exceed 35 characters
        meaning = row['Meaning'] if len(row['Meaning']) <= 35 else row['Meaning'][:32] + '[...]'
        context = row['Context'] if len(row['Context']) <= 35 else row['Context'][:32] + '[...]'
        np = row['NP'] if len(row['NP']) <= 35 else row['NP'][:32] + '[...]'
        
        formatted_output += "{:<10} {:<40} {:<40} {:<30}".format(str(row['INDEX']), meaning, context, np) + "\n"
    return formatted_output

# Create the Tkinter root window and run the application
if __name__ == "__main__":
    root = tk.Tk()
    app = PhraseMatching(root)
    root.mainloop()