import pandas as pd
import tkinter as tk
from tkinter import ttk

# Load the Excel file
#file_path = 'C:/Users/ellen/Downloads/python/db_PARTING.xlsx'
#file_path = 'C:/python projects/Parting/db_PARTING.xlsx'
file_path = '/Users/erickvillanueva/Documents/GitHub/Parting/db_PARTING.xlsx'
df = pd.read_excel(file_path)

# Group labels by category
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
        self.root.geometry("1200x600")

        # Main container setup
        main_frame = tk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=1)

        canvas = tk.Canvas(main_frame)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=1)

        scrollbar_y = ttk.Scrollbar(main_frame, orient=tk.VERTICAL, command=canvas.yview)
        scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)

        canvas.configure(yscrollcommand=scrollbar_y.set)
        canvas.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

        content_frame = tk.Frame(canvas)
        canvas.create_window((0, 0), window=content_frame, anchor="nw")

        # Storing selection widgets
        self.selections = {}

        # Layout control
        current_col = 0
        current_row = 0
        row_offset = 2
        max_columns = 6

        # Create labels and comboboxes for each group and category
        for group_name, labels in label_groups.items():
            ttk.Label(content_frame, text=group_name, font=("Arial", 12, "bold")).grid(
                row=current_row, column=current_col, columnspan=2, padx=5, pady=5, sticky="n")

            for idx, label in enumerate(labels):
                ttk.Label(content_frame, text=label, font=("Arial", 8)).grid(
                    row=row_offset + current_row + idx, column=current_col, padx=5, pady=2, sticky="e")

                # Populate dropdown values
                values = df[label].dropna().unique().tolist()
                values = [str(v) for v in values]

                if df[label].isna().any():
                    values.append('None')

                combo = ttk.Combobox(content_frame, values=['ALL'] + values, font=("Arial", 8), width=20)
                combo.grid(row=row_offset + current_row + idx, column=current_col + 1, padx=5, pady=2, sticky="w")
                combo.set('ALL')
                self.selections[label] = combo

            current_col += 2
            if current_col >= max_columns * 2:
                current_col = 0
                current_row += len(labels) + row_offset

        max_group_height = len(max(label_groups.values(), key=len)) + row_offset
        submit_btn = ttk.Button(content_frame, text="Submit", command=self.filter_data)
        submit_btn.grid(row=current_row + max_group_height + 1, column=5, columnspan=2, pady=25)

        reset_btn = ttk.Button(content_frame, text="Reset", command=self.reset_filters)
        reset_btn.grid(row=current_row + max_group_height + 1, column=7, columnspan=2, pady=25)

        self.final_result_count = ttk.Label(content_frame, text="Total Count: 0", font=("Arial", 10, "bold"))
        self.final_result_count.grid(row=current_row + max_group_height + 2, column=0, columnspan=2, pady=5)

        # Treeview for result display
        self.result_tree = ttk.Treeview(content_frame, columns=('INDEX', 'NP', 'Context', 'Meaning'), show='headings')
        self.result_tree.heading('INDEX', text='INDEX')
        self.result_tree.heading('NP', text='NP')
        self.result_tree.heading('Context', text='Context')
        self.result_tree.heading('Meaning', text='Meaning')

        for col in ['INDEX', 'NP', 'Context', 'Meaning']:
            self.result_tree.column(col, anchor='w', width=150, stretch=True)

        self.result_tree.grid(row=current_row + max_group_height + 3, column=0, columnspan=8, pady=25, sticky="nsew")

        result_scrollbar = ttk.Scrollbar(content_frame, orient=tk.VERTICAL, command=self.result_tree.yview)
        result_scrollbar.grid(row=current_row + max_group_height + 3, column=8, sticky='ns')
        self.result_tree.configure(yscrollcommand=result_scrollbar.set)

        # Configure row and column expansion
        content_frame.grid_rowconfigure(current_row + max_group_height + 3, weight=1)
        content_frame.grid_columnconfigure(0, weight=1)

        for col in range(8):
            content_frame.grid_columnconfigure(col, weight=1)

        # Apply disregard rules on selection
        for group_name, labels in label_groups.items():
            for label in labels:
                self.selections[label].bind("<<ComboboxSelected>>", lambda e: self.apply_disregard_rules())

    def apply_disregard_rules(self):
        # Clear disregard rules and reapply based on current selection
        for group_name, labels in label_groups.items():
            for label in labels:
                self.selections[label].config(state='normal')

        # Apply disregard rules conditionally as required
        pro_nominal_selected = any(self.selections[label].get() != 'ALL' for label in label_groups['PRO_Nominal'])
        if pro_nominal_selected:
            disable_groups = ['REL_Type', 'OUT_SYN', 'OUT_SEM', 'INN_Syn', 'ADJ', 'INN_Lex', 'INN_ctxt']
            for group in disable_groups:
                for label in label_groups[group]:
                    self.selections[label].config(state='disabled')

        # Example disregard rules based on specific conditions
        if self.selections['DET'].get() == 'None':
            self.selections['Q_COMPLEXITY'].set('ALL')
            self.selections['Q_COMPLEXITY'].config(state='disabled')
            self.selections['Q_INFL'].set('ALL')
            self.selections['Q_INFL'].config(state='disabled')

    def filter_data(self):
        # Filter the dataframe based on selected options
        final_filtered_df = df.copy()

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

            final_filtered_df = pd.merge(final_filtered_df, group_filtered_df, how='inner').drop_duplicates()

        # Update the result display
        for row in self.result_tree.get_children():
            self.result_tree.delete(row)

        for index, row in final_filtered_df.iterrows():
            self.result_tree.insert("", "end", values=(row['INDEX'], row['NP'], row['Context'], row['Meaning']))

        final_count = len(final_filtered_df)
        self.final_result_count.config(text=f"Total Count: {final_count}")

        # Save results to a text file
        file_path = 'parting_search.txt'  # Adjust path as needed
        headers = ['INDEX', 'NP', 'Context', 'Meaning']
        selected_options = {label: self.selections[label].get() for group_name, labels in label_groups.items() for label in labels}

        with open(file_path, 'w') as f:
            f.write("Selected Options:\n")
            for label, option in selected_options.items():
                f.write(f"{label}: {option} | ")
            f.write("\n\nResulting Rows:\n\n")

            # Write headers
            f.write("{:<10} {:<30} {:<50} {:<50}".format('INDEX', 'NP', 'Context', 'Meaning') + "\n")

            # Initialize formatted_output
            formatted_output = ""

            # Write each row of the filtered results
            for _, row in final_filtered_df[headers].iterrows():
                # Shorten fields to 32 characters and append '[...]' if they exceed 35 characters
                meaning = row['Meaning'] if len(row['Meaning']) <= 35 else row['Meaning'][:38] + '[...]'
                context = row['Context'] if len(row['Context']) <= 35 else row['Context'][:38] + '[...]'
                np = row['NP'] if len(row['NP']) <= 35 else row['NP'][:37] + '[...]'
                
                formatted_output = "{:<10} {:<30} {:<50} {:<50}".format(str(row['INDEX']), np, context, meaning) + "\n"
                f.write(formatted_output)

    def reset_filters(self):
        # Reset all filters to default values and re-enable all selections
        for label, combo in self.selections.items():
            combo.set('ALL')
            combo.config(state='normal')

        # Clear the result display in the Treeview
        for row in self.result_tree.get_children():
            self.result_tree.delete(row)

        # Reset the result count display
        self.final_result_count.config(text="Total Count: 0")

root = tk.Tk()
app = PhraseMatching(root)
root.mainloop()
