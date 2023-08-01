"This program takes all the pdfs a user selects and an excel file with a list of items to search for and searches through all the pdfs, highlighting any instance of that item and then saving the highlighted pdfs with pages removed or not removed in a folder where the original pdfs were" 

## pip install tk pymupdf pandas tqdm ##
import tkinter as tk
from tkinter import Tk, filedialog, messagebox, scrolledtext, Checkbutton
import fitz
import time
import os
import pandas as pd
import subprocess


## Method for button click 
def highlight_pdfs():
	"This method performs the main functionality of highlighting files, accessing pdf to highlight and an excel file from the user"
	#Disable user from changing preference or repressing once highlight button pressed
	comb_pages_checkbutton.config(state=tk.DISABLED)
	highlight_button.config(state = tk.DISABLED)
	
	#get file paths pdfs, start from where the exe is
	pdf_file_paths = filedialog.askopenfilenames(initialdir = os.getcwd(), title="Select PDF files to highlight", filetypes=(("PDF files","*.pdf"),))
	
	# Check for if user exits selecting pdf, resets program
	if len(pdf_file_paths) ==0:
		comb_pages_checkbutton.config(state=tk.ACTIVE)
		highlight_button.config(state = tk.ACTIVE)
		return
	
	#get excel sheet, start from where exe is
	excel_file_path = filedialog.askopenfilename(initialdir = os.getcwd(), title = "Select Excel File with items to search", filetypes=(("Excel files", "*.xlsx"),))
	
	# Checks for no excel file selected, user clicked out of selection window, resets program
	if not excel_file_path:
		comb_pages_checkbutton.config(state=tk.ACTIVE)
		highlight_button.config(state = tk.ACTIVE)
		return
	
	# Read the codes from the Excel sheet (assuming they are in the first column)
	df = pd.read_excel(excel_file_path, header=None)
	pre_codes = df.iloc[:, 0].tolist() #list of codes
	codes = pre_codes[2:] #removes first two items for formatting with sheets
	missing_codes = [True] * len(codes) #list of booleans attached to the code for finding unused codes in each PDF
	
	# Create the progress bar
	
	
	# Create a scrolledtext box to display progress
	scroll_text.configure(state='normal')
	scroll_text.delete('1.0', tk.END) #wipe at beginning of each new request
	scroll_text.configure(state = 'disabled')
	
	## Go through each pdf selected
	for i, pdf_file_path in enumerate(pdf_file_paths):	
		# Update progress bar and scroll_text box
		progress_percentage = (i+1)/len(pdf_file_paths)*100
		progress_text = f"\nProcessing PDFs ... {i+1}/{len(pdf_file_paths)} {progress_percentage:.1f}%\n\n"
		# first pass through for formatting
		if i ==0:
			progress_text = f"Processing PDFs ... {i+1}/{len(pdf_file_paths)} {progress_percentage:.1f}%\n\n"
		scroll_text.configure(state = 'normal')
		scroll_text.insert(tk.END, progress_text)
		scroll_text.see(tk.END) #scroll to bottom
		scroll_text.configure(state = 'disabled')
		root.update() #yield control to main loop so that program prints updates and doesn't stall
	
		# Open the PDF document using fitz (installed with pip install pymupdf)
		doc = fitz.open(pdf_file_path)

		# Array to store pages containing a code so only those pages are saved in the output
		pages_with_code =[]
		
		# initialize each time a new pdf is presented
		missing_codes = [True]* len(codes)

		# Iterate over pages
		for page_num in range(len(doc)):
			page = doc[page_num]
			code_found = False #flag for pages_with_code
			for i, code in enumerate(codes):
				text_instances=page.search_for(code) #finds the codes in the text
				if text_instances:
					code_found = True
					#if a certain code is found at least once then change that index of codes array to False in missing array
					missing_codes[i] = False
					#functionality for adding yellow highlight box (comes from fitz package)
					for rect in text_instances:
						highlight = page.add_highlight_annot(rect)
			
			if code_found:
			#if the page is one with at least a code, add it to this array for combed pdf
				pages_with_code.append(page_num)
		
		# no highlighted values found, dont print
		if len(pages_with_code) == 0:
			scroll_text.configure(state = 'normal')
			scroll_text.insert(tk.END, "No keyword values found in document: " + os.path.basename(pdf_file_path) +"\n")
			scroll_text.see(tk.END)
			scroll_text.configure(state = 'disabled')
			root.update()
		else:
			# Save the output in the output folder, only the select pages
			output_folder = os.path.join(os.path.dirname(pdf_file_path), "Highlighted_Outputs") #makes folder in same place it found pdf(s)
			os.makedirs(output_folder, exist_ok = True) #makes folder if not already created
			output_file_path = os.path.join(output_folder, "Highlighted_" + os.path.basename(pdf_file_path)) #adds highlighted before the name
			
			#print the missing codes from each
			for i, missing in enumerate(missing_codes):
				if missing:
					code_not_found = codes[i] #get the actual code that wasnt found, missing is boolean array
					scroll_text.configure(state = 'normal')
					scroll_text.insert(tk.END, f"{code_not_found} is not present in {os.path.basename(pdf_file_path)}\n") #prints with what pdf
					scroll_text.see(tk.END)
					scroll_text.configure(state = 'disabled')
					root.update()
			# Configurations for selecting pages, saving and closing the original
			if not comb_pages_var.get():
				#if user did not select to remove pages without a highlight save all pages
				doc.save(output_file_path)
				doc.close()
			else:
				new_doc = fitz.open()
				for page_num in pages_with_code:
					#only pages with a highlight
					new_doc.insert_pdf(doc, from_page = page_num, to_page=page_num)
					
				new_doc.save(output_file_path)
				new_doc.close()
		
	# Close progress bar
	scroll_text.configure(state = 'normal')
	#add closing message
	tabs = "\t"
	scroll_text.insert(tk.END,tabs+"\n\n\n")
	scroll_text.insert(tk.END,tabs+" _    _              _       _              _        _    \n")
	scroll_text.insert(tk.END,tabs+"| |  | | (_)        | |     | | (_)        | |     _| |_  \n")
	scroll_text.insert(tk.END,tabs+"| |__| |       __   | |__   | |       __   | |___ |_   _| \n")
	scroll_text.insert(tk.END,tabs+"|  __  | | |  /  \  |  _  \ | | | |  /  \  |  _  \  | |   \n")
	scroll_text.insert(tk.END,tabs+"| |  | | | | | () | | | | | | | | | | () | | | | |  | |   \n")
	scroll_text.insert(tk.END,tabs+"|_|  |_| |_|  \_  | |_| |_| |_| |_|  \_  | |_| |_|  |_|   \n")
	scroll_text.insert(tk.END,tabs+"                | |                    | |                \n")
	scroll_text.insert(tk.END,tabs+"               |__|                   |__|                \n")
	scroll_text.insert(tk.END,tabs+"                                                          \n")
	scroll_text.insert(tk.END,tabs+"  ____                            _            _          \n")
	scroll_text.insert(tk.END,tabs+" / ___|                          | |   ___   _| |_   ___  \n")
	scroll_text.insert(tk.END,tabs+"| |      __    ____ ____    __   | |  / _ \ |_   _| / _ \ \n")
	scroll_text.insert(tk.END,tabs+"| |     /  \  /  _ '  _ \  /  \  | | | ___|   | |  | ___| \n")
	scroll_text.insert(tk.END,tabs+"| |___ | () | | | | | | | | () | | | | |___   | |  | |___ \n")
	scroll_text.insert(tk.END,tabs+" \____| \__/  |_| |_| |_| | __/  |_|  \____|  |_|   \____|\n")
	scroll_text.insert(tk.END,tabs+"                          | |                             \n")
	scroll_text.insert(tk.END,tabs+"                          |_|                             \n")
	scroll_text.see(tk.END)
	comb_pages_checkbutton.config(state=tk.ACTIVE)
	highlight_button.config(state = tk.ACTIVE)
	scroll_text.configure(state = 'disabled')
	#Show message box confirming highlight-tion, option to open output folder
	result = messagebox.askquestion("Highlighting Complete", "PDFs saved. Open Highlighted_Outputs folder?")
	if result == "yes":
		folder_path = os.path.join(os.path.dirname(pdf_file_paths[0]), "Highlighted_Outputs")
		subprocess.Popen(f'explorer "{os.path.abspath(folder_path)}"')

# Create the main Tkinter window
root = tk.Tk()
root.title("PDF Highlighting")

# Create a button to trigger the PDF highlighting process
highlight_button = tk.Button(root, text="Highlight PDFs", command=highlight_pdfs)
highlight_button.grid(row =1, column =1, sticky = "w", padx =30, pady =5)

# Page selection checkbox setup
comb_pages_var = tk.BooleanVar()
comb_pages_var.set(True)
comb_pages_checkbutton = Checkbutton(root, text ="Dont include pages without instance of keyword in output file? (Does not modify original files)", variable = comb_pages_var, onvalue=True, offvalue=False)
comb_pages_checkbutton.grid(row =1, column =2, sticky = "w", padx = 50)

# Create a scrolled text box
scroll_text = scrolledtext.ScrolledText(root, width=75, height=30, state='disabled')
scroll_text.grid(row =2, column =1, columnspan =2, sticky = "wens", padx =5, pady=5)

root.grid_rowconfigure(2, weight =1) #expands row 2 to fill screen
root.grid_columnconfigure(1, weight =1) #expands column 1 to fill screen

# Sets the window location on the console
root.geometry("+175+175")

# Run the Tkinter event loop
root.mainloop()