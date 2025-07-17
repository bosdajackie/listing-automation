import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import threading
import pandas as pd
import textwrap
import json
from openai import OpenAI
import os
from dotenv import load_dotenv
load_dotenv()
from vehicleCompatibility import WebScraper

class ProductListingGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("3DSellers Product Listing Tool")
        self.root.geometry("800x900")
        
        # Initialize variables
        self.webscraper = None
        self.product_results = []
        self.selected_category = ""
        
        self.create_widgets()
    
    def create_widgets(self):
        # Outer frame to center content
        outer_frame = ttk.Frame(self.root)
        outer_frame.grid(row=0, column=0, padx=20, pady=20)
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

        # Center outer_frame content
        outer_frame.columnconfigure(0, weight=1)

        # Create the notebook (tab container)
        notebook = ttk.Notebook(outer_frame)
        notebook.grid(row=0, column=0, sticky="n")

        # Create tabs (frames)
        self.compatibility_tab = ttk.Frame(notebook)
        self.listing_tab = ttk.Frame(notebook)

        notebook.add(self.compatibility_tab, text="Compatibility/Specifications Checker")
        notebook.add(self.listing_tab, text="Product Listing")

        # Compatibility (main) frame
        main_frame = ttk.Frame(self.compatibility_tab, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Listing frame
        listing_frame = ttk.Frame(self.listing_tab, padding="10")
        listing_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)

        """COMPATIBILITIES TAB"""
        # Title
        title_label = ttk.Label(main_frame, text="Get Compatibility and Specifications for Part ID", 
                                font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=2, pady=(0, 20))

        # Browser Configuration
        browser_frame = ttk.LabelFrame(main_frame, text="Browser Settings", padding="10")
        browser_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(10, 10))
        
        self.headless_var = tk.BooleanVar()
        ttk.Checkbutton(browser_frame, text="Run browser in headless mode (background)", 
                        variable=self.headless_var).grid(row=0, column=0, sticky=tk.W)

        # Storefront selection
        ttk.Label(main_frame, text="Storefront:").grid(row=2, column=0, sticky=tk.W, pady=5)
        self.storefront_var = tk.StringVar()
        storefront_combo = ttk.Combobox(main_frame, textvariable=self.storefront_var, 
                                       values=["Autofirst", "Karshield", "365Hubs"],
                                       state="readonly", width=40)
        storefront_combo.grid(row=2, column=1, sticky=(tk.W, tk.E), pady=5)
        storefront_combo.set("Karshield")  # Default selection

        # Part Number Entry
        ttk.Label(main_frame, text="Part Number:").grid(row=3, column=0, sticky=tk.W, pady=5)
        self.part_number_var = tk.StringVar()
        part_number_entry = ttk.Entry(main_frame, textvariable=self.part_number_var, width=40)
        part_number_entry.grid(row=3, column=1, sticky=(tk.W, tk.E), pady=5)

        # Get Compatibility/Specs Button
        self.get_data_btn = ttk.Button(main_frame, text="Get All Product Listings", 
                                      command=self.start_webscraper)
        self.get_data_btn.grid(row=4, column=0, columnspan=2, pady=20)

        # Progress bar
        self.progress = ttk.Progressbar(main_frame, mode='indeterminate')
        self.progress.grid(row=5, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)

        # Status label
        self.status_var = tk.StringVar()
        self.status_var.set("Ready")
        status_label = ttk.Label(main_frame, textvariable=self.status_var)
        status_label.grid(row=6, column=0, columnspan=2, pady=5)

        # Product selection frame (initially hidden)
        self.product_frame = ttk.LabelFrame(main_frame, text="Product Selection", padding="10")
        for c in (0, 1): # column 0 and column 1 of self.product_frame to both expand equally when there’s extra horizontal space
            self.product_frame.columnconfigure(c, weight=1)
        self.product_frame.grid(row=7, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=10)
        self.product_frame.grid_remove()  # Hide initially

        # Specs product selection
        ttk.Label(self.product_frame, text="Select product for specifications:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.specs_product_var = tk.StringVar()
        self.specs_combo = ttk.Combobox(self.product_frame, textvariable=self.specs_product_var, 
                                        state="readonly", width=60)
        self.specs_combo.grid(row=0, column=1, sticky=(tk.W, tk.E), pady=5)

        # Compatibility product selection
        ttk.Label(self.product_frame, text="Select product for compatibility:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.compat_product_var = tk.StringVar()
        self.compat_combo = ttk.Combobox(self.product_frame, textvariable=self.compat_product_var, 
                                        state="readonly", width=60)
        self.compat_combo.grid(row=1, column=1, sticky=(tk.W, tk.E), pady=5)

        # Selected category display
        ttk.Label(self.product_frame, text="Selected Category:").grid(row=2, column=0, sticky=tk.W, pady=5)
        self.category_var = tk.StringVar()
        category_entry = ttk.Entry(self.product_frame, textvariable=self.category_var, 
                                   state="readonly", width=60)
        category_entry.grid(row=2, column=1, sticky=(tk.W, tk.E), pady=5)

        # Process selections button
        self.process_btn = ttk.Button(self.product_frame, text="Process Selections", 
                                      command=self.process_selections)
        self.process_btn.grid(row=3, column=0, columnspan=2, pady=10)

        # Results frame
        results_frame = ttk.LabelFrame(main_frame, text="Results", padding="10")
        results_frame.grid(row=8, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=10)
        
        # Configure results frame to expand
        main_frame.rowconfigure(8, weight=1)
        results_frame.columnconfigure(0, weight=1)
        results_frame.rowconfigure(0, weight=1)
        
        # Results text area
        self.results_text = scrolledtext.ScrolledText(results_frame, width=80, height=15)
        self.results_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)

        # Buttons frame
        buttons_frame = ttk.Frame(main_frame)
        buttons_frame.grid(row=9, column=0, columnspan=2, pady=10)
        
        # Clear results button
        clear_btn = ttk.Button(buttons_frame, text="Clear Results", command=self.clear_results)
        clear_btn.grid(row=0, column=0, padx=5)
        
        # Save results button
        save_btn = ttk.Button(buttons_frame, text="Save Results", command=self.save_results)
        save_btn.grid(row=0, column=1, padx=5)

        """PRODUCT LISTING TAB"""
        # Title
        title_label = ttk.Label(listing_frame, text="Get Information for Product Listing", 
                                font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=2, pady=(0, 20))

        # Selected part id
        ttk.Label(listing_frame, text="Selected Part Number:").grid(row=1, column=0, sticky=tk.W, pady=5)
        part_entry = ttk.Entry(listing_frame, textvariable=self.part_number_var, 
                               state="readonly", width=60)
        part_entry.grid(row=1, column=1, sticky=(tk.W, tk.E), pady=5)

        # Selected category display
        ttk.Label(listing_frame, text="Selected Category:").grid(row=2, column=0, sticky=tk.W, pady=5)
        category_entry = ttk.Entry(listing_frame, textvariable=self.category_var, 
                                   state="readonly", width=60)
        category_entry.grid(row=2, column=1, sticky=(tk.W, tk.E), pady=5)

        # Alternate Numbers
        ttk.Label(listing_frame, text="Alternate Numbers:").grid(row=3, column=0, sticky=(tk.W, tk.N), pady=5)
        
        # Frame for alternate numbers and instructions
        alt_frame = ttk.Frame(listing_frame)
        alt_frame.grid(row=3, column=1, sticky=(tk.W, tk.E), pady=5)
        alt_frame.columnconfigure(0, weight=1)
        
        ttk.Label(alt_frame, text="(Comma or line separated)", 
                 font=("Arial", 8), foreground="gray").grid(row=0, column=0, sticky=tk.W)
        
        self.alternate_numbers_text = scrolledtext.ScrolledText(alt_frame, height=4, width=40)
        self.alternate_numbers_text.grid(row=1, column=0, sticky=(tk.W, tk.E))

        # AI TOOLS FRAME
        ai_frame = ttk.LabelFrame(listing_frame, text="AI Generation Tools", padding="10")
        ai_frame.grid(row=4, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(10, 10))
        
        # product listing title
        self.listing_title = ttk.Entry(ai_frame, textvariable=self.part_number_var, width=80)
        self.listing_title.grid(row=0, column=0, columnspan=2, pady=10, sticky="ew")

        # generate title button
        gen_title_btn = ttk.Button(ai_frame, text="Generate Title", command=self.generate_short_desc)
        gen_title_btn.grid(row=1, column=0, columnspan=2, pady=5)

        # AI generated short desc text field
        self.short_desc_text = scrolledtext.ScrolledText(ai_frame, width=80, height=8, wrap=tk.WORD)
        self.short_desc_text.grid(row=2, column=0, columnspan=2, pady=10, sticky="ew")

        # Generate short desc button
        gen_short_btn = ttk.Button(ai_frame, text="Generate Short Desc", command=self.generate_short_desc)
        gen_short_btn.grid(row=3, column=0, columnspan=2, pady=5)

        # AI generated long desc text field
        self.long_desc_text = scrolledtext.ScrolledText(ai_frame, width=80, height=8, wrap=tk.WORD)
        self.long_desc_text.grid(row=4, column=0, columnspan=2, pady=10, sticky="ew")

        # Generate long desc button
        gen_long_btn = ttk.Button(ai_frame, text="Generate Long Desc", command=self.generate_short_desc)
        gen_long_btn.grid(row=5, column=0, columnspan=2, pady=5)

        # vehicle selection for image generation
        ttk.Label(ai_frame, text="Chosen Vehicle:").grid(row=6, column=0, sticky=tk.W, pady=5)
        self.vehicle_var = tk.StringVar()
        self.vehicle_combo = ttk.Combobox(ai_frame, textvariable=self.vehicle_var, values=[], state="readonly", width=80)
        self.vehicle_combo.grid(row=6, column=1, sticky=(tk.W, tk.E), pady=5)
        self.vehicle_combo.set("")  # Default selection

        # Generate img button
        gen_img_btn = ttk.Button(ai_frame, text="Generate Vehicle Image", command=self.generate_short_desc)
        gen_img_btn.grid(row=7, column=0, columnspan=2, pady=5)
        
        # ttk.Label(listing_frame, text="eBay Product Title:").grid(row=1, column=0, sticky=tk.W)
        # self.ebay_title_var = tk.StringVar()
        # ttk.Entry(listing_frame, textvariable=self.ebay_title_var, width=60).grid(row=1, column=1, pady=5)



    def start_webscraper(self):
        """Start the webscraper in a separate thread"""
        if not self.part_number_var.get().strip():
            messagebox.showerror("Error", "Please enter a part number")
            return
        
        # Disable the button and show progress
        self.get_data_btn.config(state='disabled')
        self.progress.start()
        self.status_var.set("Initializing webscraper...")
        
        # Start webscraper in separate thread
        thread = threading.Thread(target=self.run_webscraper)
        thread.daemon = True
        thread.start()

    def run_webscraper(self):
        """Run the webscraper and handle results"""
        try:
            # Initialize webscraper
            self.webscraper = WebScraper(
                storefront=self.storefront_var.get(),
                headless=self.headless_var.get(),
                status_callback=self.update_status
            )
            
            # Search for products
            self.update_status("Searching for products...")
            self.product_results = self.webscraper.search_products(self.part_number_var.get().strip())
            
            # Update GUI in main thread
            self.root.after(0, self.handle_product_results)
            
        except Exception as e:
            self.root.after(0, lambda: self.handle_error(str(e)))

    def handle_product_results(self):
        """Handle the product results from webscraper"""
        self.progress.stop()
        self.get_data_btn.config(state='normal')
        
        if not self.product_results:
            self.status_var.set("No products found")
            return
        
        # Populate the comboboxes
        product_options = []
        for i, product in enumerate(self.product_results):
            option = f"({i+1}) {product['part_number']} - {product['manufacturer']} - {product['category']}"
            product_options.append(option)
        
        # Add skip option for specs
        specs_options = ["Skip specifications"] + product_options
        self.specs_combo['values'] = specs_options
        self.specs_combo.set("Skip specifications")
        
        # Set compatibility options
        self.compat_combo['values'] = product_options
        if product_options:
            self.compat_combo.set(product_options[0])
            # Set initial category
            self.category_var.set(self.product_results[0]['category'])
        
        # Show the product selection frame
        self.product_frame.grid()
        self.status_var.set("Please select products for processing")
        
        # Bind category update to compatibility selection
        self.compat_combo.bind('<<ComboboxSelected>>', self.update_category)

    def update_category(self, event=None):
        """Update the category based on compatibility selection"""
        selected = self.compat_combo.get()
        if selected:
            # Extract index from selection
            try:
                index = int(selected.split(')')[0].replace('(', '')) - 1
                if 0 <= index < len(self.product_results):
                    self.category_var.set(self.product_results[index]['category'])
            except (ValueError, IndexError):
                pass

    def process_selections(self):
        """Process the selected products"""
        if not self.compat_combo.get():
            messagebox.showerror("Error", "Please select a product for compatibility")
            return
        
        # Get selected indices
        specs_selection = self.specs_combo.get()
        compat_selection = self.compat_combo.get()
        
        specs_index = None
        if specs_selection != "Skip specifications":
            try:
                specs_index = int(specs_selection.split(')')[0].replace('(', '')) - 1
            except (ValueError, IndexError):
                pass
        
        try:
            compat_index = int(compat_selection.split(')')[0].replace('(', '')) - 1
        except (ValueError, IndexError):
            messagebox.showerror("Error", "Invalid compatibility selection")
            return
        
        # Disable controls and show progress
        self.process_btn.config(state='disabled')
        self.progress.start()
        self.status_var.set("Processing selections...")
        
        # Start processing in separate thread
        thread = threading.Thread(target=self.run_processing, args=(specs_index, compat_index))
        thread.daemon = True
        thread.start()

    def run_processing(self, specs_index, compat_index):
        """Run the processing in background thread"""
        try:
            # Process specifications if requested
            if specs_index is not None:
                self.update_status("Getting specifications...")
                self.webscraper.get_specifications(specs_index)
            
            # Process compatibility
            self.update_status("Getting compatibility information...")
            results = self.webscraper.get_compatibility(compat_index)
            
            # Update GUI with results
            self.root.after(0, lambda: self.handle_processing_results(results))
            
        except Exception as e:
            self.root.after(0, lambda: self.handle_error(str(e)))

    def handle_processing_results(self, results):
        """Handle the processing results"""
        self.progress.stop()
        self.process_btn.config(state='normal')
        self.status_var.set("Processing completed")
        
        # Display results
        self.results_text.delete(1.0, tk.END)
        self.results_text.insert(tk.END, results)
        
        # Refresh vehicle dropdown values for image generation
        vehicle_list = self.get_vehicles()
        if vehicle_list:
            self.vehicle_combo['values'] = vehicle_list
            self.vehicle_combo.set("")  # Reset selection
        
        # Close webscraper
        if self.webscraper:
            self.webscraper.close()        

    def handle_error(self, error_msg):
        """Handle errors from webscraper"""
        self.progress.stop()
        self.get_data_btn.config(state='normal')
        self.process_btn.config(state='normal')
        self.status_var.set("Error occurred")
        messagebox.showerror("Error", f"An error occurred: {error_msg}")
        
        if self.webscraper:
            self.webscraper.close()

    def update_status(self, message):
        """Update status message (thread-safe)"""
        self.root.after(0, lambda: self.status_var.set(message))

    def clear_results(self):
        """Clear the results text area"""
        self.results_text.delete(1.0, tk.END)

    def save_results(self):
        """Save results to file"""
        if not self.results_text.get(1.0, tk.END).strip():
            messagebox.showwarning("Warning", "No results to save")
            return
        
        file_path = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")]
        )
        
        if file_path:
            try:
                with open(file_path, 'w', encoding='utf-8') as f:
                    f.write(self.results_text.get(1.0, tk.END))
                messagebox.showinfo("Success", f"Results saved to {file_path}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to save file: {str(e)}")

    # get alternate numbers from gui field and return as array
    def get_alternate_numbers(self):
        """Parse alternate numbers from text area"""
        text = self.alternate_numbers_text.get("1.0", tk.END).strip()
        if not text:
            return []
        
        # Split by comma or newline
        numbers = []
        for line in text.split('\n'):
            if ',' in line:
                numbers.extend([num.strip() for num in line.split(',') if num.strip()])
            else:
                if line.strip():
                    numbers.append(line.strip())
        
        return numbers

    # Read compatibility.xlsx and generate short desc
    def generate_short_desc(self):
        excel_path = self.webscraper.compatibility_excel_path
        if not os.path.exists(excel_path):
            messagebox.showerror("Error", "compatibility.xlsx not found.")
            return
        
        # read excel sheet
        try:
            df = pd.read_excel(excel_path)
        except Exception as e:
            messagebox.showerror("Error", f"Could not read Excel: {e}")
            return
        
        # build short desc manually
        rows = []
        for _, r in df.iterrows():
            rows.append(f"{r['Make']} {r['Model']} {r['Year']} {r['Position']} {r['Engine']}")
        plain = "Fits " + ", ".join(rows) + "."
        vehicle_list = json.dumps(rows)
        category = self.category_var.get().strip()

        # get GPT to write short desc
        if "OPENAI_API_KEY" in os.environ:
            client = OpenAI()
            # models = client.models.list()
            # for model in models.data:
            #     print(model.id)
            rules = (
                "You are generating a short e-commerce description for a wheel hub and bearing assembly based on vehicle compatibility data.\n\n"
                "Follow these exact instructions:\n"
                "- Each vehicle line must begin with space-separated years (e.g., 2014 2015 2016).\n"
                "- Then list Make, Model, and side or position (e.g., Front Right).\n"
                "- If the combined years + Make/Model/Position exceeds 65 characters, split the years across multiple lines. Each line must still repeat the full Make, Model, and Position.\n"
                "- No hyphens or en dashes in year ranges. List years explicitly.\n"
                "- One complete fitment per line, Each sentence must end with a real newline (\n) NOT TAB, No two sentences on the same line, Never wrap within a sentence — only between sentences!!!! VERY IMPORTANT\n"
                "- MOST IMPORTANT PART: Make sure that after every sentence (after every period) a NEW LINE STARTS. NO TABS WHATSOEVER, thanks"
                "- Do NOT guess drive type, lug count, or features. FACT CHECK !!!!!!\n"
                "- Make sure that each sentence that you generate is a new line.\n"
                "- After listing all vehicles, include exactly 4 final lines:\n"
                "  1. Line describing which sides/positions it fits (e.g., 'Front Left Right...').\n"
                "  2. Line describing what the part includes (e.g., 'Includes hub, bearing...').\n"
                "  3. Line summarizing verified physical or trim details (e.g., 'Fits AWD 5 Lug 5 Bolt 5 Stud.').\n"
                "  4. Final line: Compatibility summary, e.g., 'Compatible with 2 and 4 Door Compact Luxury Models'. This is the only line allowed to use the word 'Compatible'. FACT CHECK !!!\n"
                "Example Input:\n"
                "BMW 228I 2014-2016 Front RWD, 12mm Bolt Mounting Dimension\n"
                "BMW 320I 2012-2016 Front RWD, 12mm Bolt Mounting Dimension\n"
                "BMW 328D 2014-2016	Front RWD, 12mm Bolt Mounting Dimension\n"
                "BMW 328I 2013-2016	Front RWD, 12mm Bolt Mounting Dimension\n"
                "BMW 335I 2014-2015	Front RWD, 12mm Bolt Mounting Dimension\n"
                "BMW 340I 2016 Front 340i Base Model, 12mm Bolt Mounting Dimension\n"
                "BMW 428I 2014-2016 Front RWD, 12mm Bolt Mounting Dimension\n"
                "BMW 435I 2014-2016 Front RWD, 12mm Bolt Mounting Dimension\n"
                "BMW ACTIVEHYBRID 3	2013-2015 Front 12mm Bolt Mounting Dimension\n"
                "BMW M235I 2014-2016 Front RWD, 12mm Bolt Mounting Dimension\n"

                "Expected Output:\n"
                "2014 2015 2016 BMW 228i RWD Front.\n"
                "2012 2013 2014 2015 2016 BMW 320i RWD Front.\n"
                "2014 2015 2016 BMW 328d RWD Front.\n"
                "2013 2014 2015 2016 BMW 328i RWD Front.\n"
                "2014 2015 BMW 335i RWD Front.\n"
                "2016 BMW 340i Base RWD Front.\n"
                "2014 2015 2016 BMW 428i RWD Front.\n"
                "2014 2015 2016 BMW 435i RWD Front.\n"
                "2013 2014 2015 BMW ActiveHybrid 3 Front.\n"
                "2014 2015 2016 BMW M235i RWD Front.\n"
                "Front Left Right, Front Left Side Right Side.\n"
                "Fits Base, Luxury, Sport, M Sport, and M trims.\n"
                "Wheel hub assembly with 12mm bolt mounting dimension.\n"
                "Fits RWD 5 Lug 5 Bolt 5 Stud.\n"
                "Compatible with 2 and 4 Door Compact Luxury Models.\n"
                "BMW 328I 2013-2016 Front RWD, 12mm Bolt Mounting Dimension\n"
                "BMW 340I 2016 Front 340i Base Model, 12mm Bolt Mounting Dimension\n"
            )

            prompt = (
                f"You are generating a short e-commerce description for a {category}.\n"
                f"Here is the vehicle data:\n{vehicle_list}\n"
                f"Use the formatting rules I gave you earlier."
            )

            try:
                resp = client.responses.create(
                    model="gpt-4o",
                    input = [{
                            "role": "developer",
                            "content": rules
                        },
                        {
                            "role": "user",
                            "content": prompt
                        }
                    ],
                    temperature=0,
                )
                # desc = resp.output_text
                desc = resp.output[0].content[0].text
            except Exception as e:
                # fallback to manual short desc
                print("OpenAI API failed:", e)
                desc = plain
        else:
            desc = plain

        # insert into gui field
        self.short_desc_text.delete("1.0", tk.END)
        self.short_desc_text.insert(tk.END, textwrap.fill(desc, 100))
    
    def get_vehicles(self):
        if not self.webscraper or not hasattr(self.webscraper, 'compatibility_excel_path'):
            return []
        
        excel_path = self.webscraper.compatibility_excel_path
        if not os.path.exists(excel_path):
            messagebox.showerror("Error", "compatibility.xlsx not found.")
            return []
        
        # read excel sheet
        try:
            df = pd.read_excel(excel_path)
        except Exception as e:
            messagebox.showerror("Error", f"Could not read Excel: {e}")
            return []
        
        # turn vehicles list into array
        vehicles = []
        for _, row in df.iterrows():
            year_range = row['Year']
            end_year = year_range
            if '-' in year_range:
                end_year = year_range.split('-')[1].strip()
            vehicles.append(f"{row['Make']} {row['Model']} {end_year}")
        
        return vehicles

if __name__ == "__main__":
    root = tk.Tk()
    app = ProductListingGUI(root)
    root.mainloop()