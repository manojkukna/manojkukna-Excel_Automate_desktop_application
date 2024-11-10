import tkinter as tk
from operator import index
from tkinter import ttk
from tkinter import messagebox
import Excel_Automate as va



# Initialize the main window
root = tk.Tk()
root.title("EXCEL AUTOMATION")
root.geometry("800x600")

# Create a main frame to hold the sidebar and content area
main_frame = tk.Frame(root)
main_frame.pack(fill="both", expand=True)

# Create a sidebar frame
sidebar = tk.Frame(main_frame, width=150, bg="black")
sidebar.pack(side="left", fill="y")

# Create a content frame for the main area
content_frame = tk.Frame(main_frame, bg="lightgray")
content_frame.pack(side="right", fill="both", expand=True)


def show_content(text):
    """Clears the content area and displays new content text."""
    for widget in content_frame.winfo_children():
        widget.destroy()
    label = tk.Label(content_frame, text=text, font=("Comic Sans MS", 20, "bold"), bg="white")
    label.pack(pady=20)


def Vlookup_Advanced_button_click():
    """Creates the layout for VLOOKUP Advanced with multiple sections."""
    # Clear content frame
    for widget in content_frame.winfo_children():
        widget.destroy()
    hedar_font = "Arial", 20, "bold"
    hedar_tite_font = "Arial", 15, "bold"
    inut_uset_font = "Arial", 10,"bold"

    title_label = tk.Label(content_frame, text="Vlookup Advanced", font=(hedar_font), bg="white")
    title_label.pack(pady=10)

    # Section {1}: Selected Table
    section1_frame = tk.LabelFrame(content_frame, text="{1} Selected Table", font=hedar_tite_font)
    section1_frame.pack(fill="x", padx=20, pady=5)

    tk.Label(section1_frame, text="Active Workbook Name:",font=(inut_uset_font)).grid(row=0, column=0, padx=5, pady=5)
    active_workbook_name_entry = tk.Entry(section1_frame, width=30, font=(inut_uset_font))  # Larger width and font size
    active_workbook_name_entry.insert(0, va.active_Workbook_name())  # Default value
    active_workbook_name_entry.grid(row=0, column=1, padx=5, pady=5)


    # Combobox for selecting Sheet Name from a list
    tk.Label(section1_frame, text="Selected Table Array Sheet Name:",font=(inut_uset_font)).grid(row=1, column=0, padx=5, pady=5)
    Selected_Table_Array_Sheet_Name = ttk.Combobox(section1_frame, values=va.sheet_names_list(), width=27,font=(inut_uset_font))
    Selected_Table_Array_Sheet_Name.set(va.sheet_names_list()[0])  # Set default value to Sheet1
    Selected_Table_Array_Sheet_Name.grid(row=1, column=1, padx=5, pady=5)

    tk.Label(section1_frame, text="Selected Table Array Range:",font=(inut_uset_font)).grid(row=2, column=0, padx=5, pady=5)
    selected_range_entry = tk.Entry(section1_frame, width=30, font=(inut_uset_font))  # Larger width and font size
    selected_range_entry.insert(0, "A1:Y51")  # Default value
    selected_range_entry.grid(row=2, column=1, padx=5, pady=5)

    # Section {2}: Selected VLOOKUP Columns and Row
    section2_frame = tk.LabelFrame(content_frame, text="{2} Selected VLOOKUP Columns and Row",font=hedar_tite_font)
    section2_frame.pack(fill="x", padx=20, pady=5)

    # Combobox for selecting Sheet Name from a list
    tk.Label(section2_frame, text="Lookup Value Sheets Name:",font=(inut_uset_font)).grid(row=0, column=0, padx=5, pady=5)
    lookup_sheet_name_entry_combobox = ttk.Combobox(section2_frame, values=va.sheet_names_list(), width=27,font=(inut_uset_font))
    lookup_sheet_name_entry_combobox.set(va.sheet_names_list()[0])  # Set default value to Sheet1
    lookup_sheet_name_entry_combobox.grid(row=0, column=1, padx=5, pady=5)


    tk.Label(section2_frame, text="Lookup Value Sheets Columns:",font=(inut_uset_font)).grid(row=1, column=0, padx=5, pady=5)
    lookup_column_entry = tk.Entry(section2_frame, width=30, font=(inut_uset_font))  # Larger width and font size
    lookup_column_entry.insert(0, "A1:Y1")  # Default value
    lookup_column_entry.grid(row=1, column=1, padx=5, pady=5)

    tk.Label(section2_frame, text="Lookup Value Sheet Row:",font=(inut_uset_font)).grid(row=2, column=0, padx=5, pady=5)
    lookup_row_entry = tk.Entry(section2_frame, width=30, font=(inut_uset_font))  # Larger width and font size
    lookup_row_entry.insert(0, "A1:A13")  # Default value
    lookup_row_entry.grid(row=2, column=1, padx=5, pady=5)

    # Section {3}: Final Output
    section3_frame = tk.LabelFrame(content_frame, text="{3} Final Output", font=hedar_tite_font)
    section3_frame.pack(fill="x", padx=20, pady=5)


    # Combobox for selecting Sheet Name from a list
    tk.Label(section3_frame, text="Final_Output_Sheet_Name:",font=(inut_uset_font)).grid(row=0, column=0, padx=5, pady=5)
    final_output_sheet_name_entry_combobox = ttk.Combobox(section3_frame, values=va.sheet_names_list(), width=27,font=(inut_uset_font))
    final_output_sheet_name_entry_combobox.set(va.sheet_names_list()[0])  # Set default value to Sheet1
    final_output_sheet_name_entry_combobox.grid(row=0, column=1, padx=5, pady=5)


    tk.Label(section3_frame, text="Final Output Cell:",font=(inut_uset_font)).grid(row=1, column=0, padx=5, pady=5)
    output_cell_entry = tk.Entry(section3_frame, width=30, font=(inut_uset_font))  # Larger width and font size
    output_cell_entry.insert(0, "A1")  # Default value
    output_cell_entry.grid(row=1, column=1, padx=5, pady=5)

    # Submit Button
    submit_button = tk.Button(content_frame, text="Submit",
                              command=lambda: display_input(active_workbook_name_entry,Selected_Table_Array_Sheet_Name,
                                                            selected_range_entry,lookup_sheet_name_entry_combobox,
                                                            lookup_column_entry, lookup_row_entry,
                                                            final_output_sheet_name_entry_combobox, output_cell_entry))
    submit_button.pack(pady=10)


def display_input(active_workbook, selected_sheet, selected_range, lookup_sheet, lookup_column, lookup_row,
                  output_sheet, output_cell):
    """Displays all inputs in a message box."""
    message = (
        f"Active Workbook Name: {active_workbook.get()}\n"
        f"Selected Table Array Sheet Name: {selected_sheet.get()}\n"
        f"Selected Table Array Range: {selected_range.get()}\n"
        f"Lookup Value Sheets Name: {lookup_sheet.get()}\n"
        f"Lookup Value Sheets Columns: {lookup_column.get()}\n"
        f"Lookup Value Sheet Row: {lookup_row.get()}\n"
        f"Final Output Sheet Name: {output_sheet.get()}\n"
        f"Final Output Cell: {output_cell.get()}"
    )
    # messagebox.showinfo("Entered Values", message)
    #



    selected_table_array = va.sheets_tabal_selet(sheet=selected_sheet.get(),
                                                 range=selected_range.get())

    print( selected_table_array)

    data_vlookup = va.df_vlookup(df=selected_table_array,
                                 vlookup_sheets=lookup_sheet.get(),
                                 filtered_columns=lookup_column.get(),
                                 filtered_row=lookup_row.get())
    print(data_vlookup)


    final_output_cell = va.final_output_cell(Sheet_names=output_sheet.get(), df=data_vlookup,
                                             range=output_cell.get())

    # copy_paste_code_to_excel----------------------------------------------------------

    def copy_code_lookup_value(lookup_row):
        lookup_value = f'${lookup_row.get()[0]}{int(lookup_row.get()[1]) + 1}'
        return lookup_value


    copy_code_lookup_value(lookup_row)


    def copy_code_table_array(selected_range):
        table_array_start_cell, table_array_end_cell = selected_range.split(':')
        table_array_start_cell_column = va.split_cell_column(cell_column=table_array_start_cell)[0]
        table_array_start_cell_row = int(va.split_cell_column(cell_column=table_array_start_cell)[1])
        table_array_start_cell = f'${table_array_start_cell_column}${table_array_start_cell_row}'

        table_array_end_cell_column = va.split_cell_column(cell_column=table_array_end_cell)[0]
        table_array_end_cell_row = int(va.split_cell_column(cell_column=table_array_end_cell)[1])
        table_array_end_cell = f'${table_array_end_cell_column}${table_array_end_cell_row}'

        table_array = f'{selected_sheet.get()}!{table_array_start_cell}:{table_array_end_cell}'
        return  table_array

    copy_code_table_array(selected_range=selected_range.get())

    def MATCH_lookup_value(lookup_value_columns):
        # Split the string into start and end cell references
        start_cell, end_cell = lookup_value_columns.split(':')
        column = va.split_cell_column(cell_column=start_cell)[0]
        row = int(va.split_cell_column(cell_column=start_cell)[1])

        column_letters_list = va.column_letters_list()
        column_letters_list_index = column_letters_list.index(column)

        MATCH_lookup_value = f'{lookup_sheet.get()}!{column_letters_list[column_letters_list_index + 1]}${row}'
        return MATCH_lookup_value





    MATCH_lookup_value(lookup_value_columns=lookup_column.get())

    def MATCH_lookup_array(lookup_value_columns):
        lookup_value_start_cell, lookup_value_end_cell = lookup_value_columns.split(':')  # lookup_value_columns

        lookup_value_start_column = va.split_cell_column(cell_column=lookup_value_start_cell)[0]
        lookup_value_start_row = int(va.split_cell_column(cell_column=lookup_value_start_cell)[1])
        lookup_value_start_cell = f'${lookup_value_start_column}${lookup_value_start_row}'

        lookup_value_end_cell_column = va.split_cell_column(cell_column=lookup_value_end_cell)[0]
        lookup_value_end_cell_row = int(va.split_cell_column(cell_column=lookup_value_end_cell)[1])
        table_array_end_cell = f'${lookup_value_end_cell_column}${lookup_value_end_cell_row}'

        MATCH_rang = f'{lookup_sheet.get()}!{lookup_value_start_cell}:{table_array_end_cell}'
        return MATCH_rang


    MATCH_lookup_array(lookup_value_columns=lookup_column.get())

    lookup_value = copy_code_lookup_value(lookup_row=lookup_row)
    table_array = copy_code_table_array(selected_range=selected_range.get())
    MATCH_lookup_valu = MATCH_lookup_value(lookup_value_columns=lookup_column.get())
    MATCH_lookup_arra = MATCH_lookup_array(lookup_value_columns=lookup_column.get())

    # st.code(f'=VLOOKUP($A2,Sheet1!$A$1:$Y$30,MATCH(Sheet2!B$1,Sheet2!$A$1:$Y$1,0),0)', language='excel')
    # st.code(f'=VLOOKUP(lookup_value,selected_sheet!selected_range,MATCH(Sheet2!B$1,Sheet2!$A$1:$Y$1,0),0)', language='excel')

    code = f"=VLOOKUP({lookup_value},{table_array},MATCH({MATCH_lookup_valu},{MATCH_lookup_arra},0),0)"

    print('lookup_value', copy_code_lookup_value(lookup_row=lookup_row))
    print('table_array', {copy_code_table_array(selected_range=selected_range.get())})
    print('MATCH_lookup_valu', MATCH_lookup_value(lookup_value_columns=lookup_column.get()))
    print('MATCH_lookup_arra', MATCH_lookup_array(lookup_value_columns=lookup_column.get()))
    print('code:' , code)

    # messagebox.showinfo("Entered Values",code)
    import pyperclip
    #
    #
    # Function to copy text to clipboard
    def copy_text(text):
        """Copy the provided text to the clipboard."""
        pyperclip.copy(text)
        messagebox.showinfo("Copied", "Text has been copied to the clipboard!")

    def show_message_with_copy_button(text):
        """Show a message in the content area and add a 'Copy' button."""
        for widget in content_frame.winfo_children():
            widget.destroy()

        # Display the message
        label = tk.Label(content_frame, text=text, font=("Arial", 12), bg="white", wraplength=600)
        label.pack(pady=10)

        # Copy button
        copy_button = tk.Button(content_frame, text="Copy Text", command=lambda: copy_text(text))
        copy_button.pack(pady=5)

    show_message_with_copy_button(text=code)




# Sidebar button functions
def home_button_click():
    show_content("EXCEL AUTOMATION")





# Create sidebar buttons
btn1 = tk.Button(sidebar, text="Home", command=home_button_click)
btn1.pack(fill="x", pady=5, padx=10)

btn2 = tk.Button(sidebar, text="Vlookup Advanced", command=Vlookup_Advanced_button_click)
btn2.pack(fill="x", pady=5, padx=10)

# Show the Home page by default
home_button_click()
Vlookup_Advanced_button_click()
# Run the application
root.mainloop()
