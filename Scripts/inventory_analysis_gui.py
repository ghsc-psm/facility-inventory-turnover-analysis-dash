from tkinter import *
from tkinter import ttk
from tkinter import filedialog
import pandas as pd 
import os
from inventory_analysis import InventoryRates
import configparser
from os import path

window = Tk()
window.title("Inventory Analysis")
window.geometry("800x700")


import logging
logging.basicConfig(filename=os.getcwd()+'/../Log/inventory_analysis.log',level=logging.DEBUG,format='%(asctime)s:%(levelname)s:%(message)s')

class IAGui:
	def __init__(self, window):
		self.config = IntVar(value=0)
		# Radiobutton(window,text="Use saved settings",variable=self.config,value=1, state=DISABLED).grid(row=1,column=2,sticky=W)
		# Radiobutton(window,text="New configurations",variable=self.config,value=0, state=DISABLED).grid(row=1,column=3,sticky=W)

		self.label_file_explorer = Label(window,text = "Select an inventory data file to begin",fg = "blue")
		self.label_file_explorer.grid(row=2,column=2)

		self.analysis = None

		file_button = Button(window,
						text = "Browse Files",
						command = self.browseFiles)
		file_button.grid(row=3,column=2)

	def browseFiles(self):
		
		if self.config.get() == 1:
			basepath = path.abspath("")

			config_obj = configparser.ConfigParser()
			config_obj.read(path.join(basepath,'configfile.ini'))

			col_names = config_obj['column headers']

			self.min_value = StringVar(value=col_names['min_value'])
			self.max_value = StringVar(value=col_names['max_value'])
			self.del_freq = StringVar(value=col_names['del_freq'])
			self.data_type = StringVar(value=col_names['data_type'])
			self.product = StringVar(value=col_names['product'])
			self.product1 = StringVar(value=col_names['product1'])
			self.product2  = StringVar(value=col_names['product2'])
			self.facility = StringVar(value=col_names['facility'])
			self.facility1 = StringVar(value=col_names['facility1'])
			self.facility2 = StringVar(value=col_names['facility2'])
			self.date_field = StringVar(value=col_names['date_field'])
			self.issue_col = StringVar(value=col_names['issue_col'])
			self.soh_col = StringVar(value=col_names['soh_col'])
			# self.stockout_col = StringVar(value=col_names['stockout_col'])

			other = config_obj['other']

			self.rates_window = IntVar(value=int(other['rates_window']))
			self.issue_0 = IntVar(value=int(other['issue_0']))
			self.soh_0 = IntVar(value=int(other['soh_0']))
			# self.stockout_0 = IntVar(value=int(other['stockout_0']))

			self.filename = ""
			self.column_names = ["N/A"]
			self.sheetname = StringVar()

		else:

			self.min_value = StringVar()
			self.max_value = StringVar()
			self.del_freq = StringVar()
			self.data_type = StringVar()
			self.product = StringVar()
			self.product1 = StringVar()
			self.product2  = StringVar()
			self.facility = StringVar()
			self.facility1 = StringVar()
			self.facility2 = StringVar()
			self.date_field = StringVar()
			self.rates_window = IntVar(value=12)
			self.issue_col = StringVar()
			self.issue_0 = IntVar(value=0)
			self.soh_col = StringVar()
			self.soh_0 = IntVar(value=0)
			# self.stockout_col = StringVar()
			# self.stockout_0 = IntVar(value=0)
			self.congif = IntVar(value=0)
			self.filename = ""
			self.column_names = ["N/A"]
			self.sheetname = StringVar()

		self.filename = filedialog.askopenfilename(initialdir = "/",title = "Select a File", filetypes = (("Excel","*.xlsx"),("CSV","*.csv")))

		if self.filename:

			self.label_file_explorer.configure(text="File Opened: "+self.filename.split("/")[-1])
			logging.info('File selected: {}'.format(self.filename))

			if str(self.filename).endswith('.xlsx'):
				#wb = openpyxl.load_workbook(self.filename)
				#print(wb.sheetnames)
				wb = pd.ExcelFile(self.filename)
				print(wb.sheet_names)
				#if len(wb.sheet_names) == 1:
				self.sheetname.set(wb.sheet_names[0])
				# else:
				# 	Label(window, text="Sheet").grid(row=1, column=1,sticky=E)
				# 	ttk.Combobox(window, textvariable=self.sheetname, values=wb.sheet_names,justify=LEFT).grid(row=1,column=2, padx=(0,5))
				# 	time.sleep(5)
				self.column_names = [x.strip() for x in list(pd.read_excel(self.filename,sheet_name = self.sheetname.get(),nrows=1).columns)]
			elif str(self.filename).endswith('.csv'):
				self.column_names = [x.strip() for x in list(pd.read_csv(self.filename,nrows=1).columns)]

			minlabel = Label(window, text="Min Value")
			minlabel.grid(row=4, column=1, sticky=NE)
			maxlabel = Label(window, text="Max Value")
			maxlabel.grid(row=4, column=3, sticky=NE)
			delfreq = Label(window, text="Delivery Frequency")
			delfreq.grid(row=5, column=1, sticky=NE)
			dataTypeLabel = Label(window, text="Data Type")
			dataTypeLabel.grid(row=6, column=1, sticky=NE)
			Label(window, text="Product Name Field").grid(row=7, column=1,sticky=E)
			Label(window, text="(Additional Product Info)").grid(row=8,column=1,sticky=E)
			Label(window, text="(Additional Product Info)").grid(row=9,column=1,sticky=NE)
			Label(window, text="Facility Name Field").grid(row=10, column=1, sticky=E)
			Label(window, text="Region Name Field").grid(row=11,column=1,sticky=E)
			Label(window, text="Select State/District/Province").grid(row=11,column=3,sticky=E)
			Label(window, text="(Additional Facility Info)").grid(row=12,column=1,sticky=NE)
			Label(window, text="Date Field").grid(row=13, column=1, sticky=E)
			Label(window, text="Consumption Field").grid(row=14, column=1, sticky=E)
			Label(window, text="Stock on Hand Field").grid(row=15, column=1, sticky=E)
			# Label(window, text="Stockout Days Field").grid(row=14, column=1, sticky=NE)
			Label(window, text="Time Window for Rates").grid(row=16, column=1, sticky=E)

			#Label(window, text=None).grid(row=11, column=1, sticky=E)
			min_entry = ttk.Combobox(window, textvariable=self.min_value, values=[1,2,3,4,5,6],
				  justify=LEFT)                              
			min_entry.grid(row=4,column=2, padx=(0,5),pady=(0,5))
			max_entry = ttk.Combobox(window, textvariable=self.max_value, values=[1,2,3,4,5,6],
				  justify=LEFT)
			max_entry.grid(row=4,column=4, padx=(0,5),pady=(0,5))
			freq_entry = ttk.Combobox(window, textvariable=self.del_freq, values=['Monthly','Bimonthly','Quarterly'],
				  justify=LEFT)
			freq_entry.grid(row=5,column=2, padx=(0,5),pady=(0,5                                                                                                                                                                                                                                                                    ))
			data_type_entry = ttk.Combobox(window, textvariable=self.data_type, values=['transactional','cumulative'],
				  justify=LEFT)
			data_type_entry.grid(row=6,column=2, padx=(0,5),pady=(0,50))
			#print(self.column_names)
			ttk.Combobox(window, textvariable=self.product, values=self.column_names,
				  justify=LEFT).grid(row=7,column=2, padx=(0,5))
			ttk.Combobox(window, textvariable=self.product1, values=self.column_names,
				  state = "disabled",justify=LEFT).grid(row=8,column=2, padx=(40,0))
			ttk.Combobox(window, textvariable=self.product2, values=self.column_names,
				  state = "disabled",justify=LEFT).grid(row=9,column=2, padx=(40,0),pady=(0,50))
			ttk.Combobox(window, textvariable=self.facility, values=self.column_names,
				  justify=LEFT).grid(row=10,column=2, padx=(0,5))
			ttk.Combobox(window, textvariable=self.facility1, values=self.column_names,
				  justify=LEFT).grid(row=11,column=2, padx=(40,0))
			ttk.Combobox(window, textvariable=self.facility2, values=self.column_names,
				  state = "disabled",justify=LEFT).grid(row=12,column=2, padx=(40,0),pady=(0,50))
			ttk.Combobox(window, textvariable=self.date_field, values=self.column_names,
				  justify=LEFT).grid(row=13,column=2, padx=(0,5))
			ttk.Combobox(window, textvariable=self.issue_col, values=self.column_names,
				  justify=LEFT).grid(row=14,column=2, padx=(0,5))
			Radiobutton(window,text="Blank fields = 0",variable=self.issue_0,value=1).grid(row=14,column=3,sticky=W,padx=(0,5))
			Radiobutton(window,text="Blank fields are missing",variable=self.issue_0,value=0).grid(row=14,column=4,sticky=W,padx=(0,5))
			ttk.Combobox(window, textvariable=self.soh_col, values=self.column_names,
				  justify=LEFT).grid(row=15,column=2, padx=(0,5))
			Radiobutton(window,text="Blank fields = 0",variable=self.soh_0,value=1).grid(row=15,column=3,sticky=W,padx=(0,5))
			Radiobutton(window,text="Blank fields are missing",variable=self.soh_0,value=0).grid(row=15,column=4,sticky=W,padx=(0,5))
			# ttk.Combobox(window, textvariable=self.stockout_col, values=self.column_names,
				#   justify=LEFT).grid(row=14,column=2, padx=(0,5),pady=(0,50))
			# Radiobutton(window,text="Blank fields = 0",variable=self.stockout_0,value=1).grid(row=14,column=3,sticky=W,padx=(0,5),pady=(0,50))
			# Radiobutton(window,text="Blank fields are missing",variable=self.stockout_0,value=0).grid(row=14,column=4,sticky=W,padx=(0,5),pady=(0,50))
			ttk.Combobox(window, textvariable=self.rates_window, values=list(range(6,25)),
				  justify=LEFT).grid(row=16,column=2, padx=(0,5))

			run_button = Button(window,
								text="Run Analysis",
								command = self.runIA)
			run_button.grid(row=17,column=2)
	
	def runIA(self):
		print("> initiating")
		products = list(filter(None,[self.product.get(),self.product1.get(),self.product2.get()]))
		facilities = list(filter(None,[self.facility.get(),self.facility1.get(),self.facility2.get()]))
		i = InventoryRates(inv_type = self.data_type.get(), data_path = self.filename, date_col = self.date_field.get(), issue_col = self.issue_col.get(), soh_col = self.soh_col.get(), window= self.rates_window.get(),agg_cols=facilities+products,fac_cols = facilities, prod_cols=products, del_freq=self.del_freq.get(), min_value=self.min_value.get(), max_value=self.max_value.get())
		Label(window, text='Calculating Rates').grid(row=18,column=2)
		logging.debug('Product columns: {}'.format(products))
		logging.debug('Facility columns: {}'.format(facilities))
		logging.info('Inventory rates object instantiated')
		print("> cleaning data")
		i.clean_text()
		print("> adjusting for transactional data (if applicable)")
		i.transactional()
		print("> begin rolling rates")
		i.rolling_rates()
		print("> end rolling rates")
		print("> begin stock status")
		i.stock_status()
		print("> end stock status")
		print("> begin subsetting raw data")
		i.subset_raw()
		print("> end subsetting raw data")
		i.set_controls()
		i.format_columns()
		print("> end format cols")
		i.create_ref_tables()
		print("> end create ref tables")
		self.analysis = i
		#temp code
		# self.analysis.raw_data['Facility Name'] = (self.analysis.raw_data['Facility Name']).apply(lambda x: str(x).strip().upper())
		# self.analysis.raw_data['Product Name'] = (self.analysis.raw_data['Product Name']).apply(lambda x: str(x).strip().upper())
		# self.analysis.raw_data['Date'] = pd.to_datetime(self.analysis.raw_data['Date'])
		# self.analysis.raw_data = self.analysis.rates_data.merge(self.analysis.raw_data, left_on=['Facility_id','Product_id','Date'],
		# 												  right_on=['Facility Name','Product Name','Date'],how='left', indicator=True)
		# print(self.analysis.raw_data.head,self.analysis.raw_data.shape)
		# self.analysis.raw_data = self.analysis.raw_data[self.analysis.raw_data['_merge'] == 'both']
		# print(self.analysis.raw_data.shape)

		#write files
		print('> writing output')
		self.analysis.raw_data.to_csv('raw_data.txt', index=None, sep=' ', mode='w')
		self.analysis.rates_data.to_csv('rates_data.txt', index=None, sep=' ', mode='w')
		self.analysis.stock_data.to_csv('stock_data.txt', index=None, sep=' ', mode='w')
		self.analysis.prod_table.to_csv('product_list.txt', index=None, sep=' ', mode='w')
		self.analysis.fac_table.to_csv('facility_list.txt', index=None, sep=' ', mode='w')
		self.analysis.plot_data.to_csv('plot_data.txt', index=None, sep=' ', mode='w')
		self.analysis.dates_table.to_csv('dates_table.txt', index=None, sep=' ', mode='w')
		self.analysis.control_table.to_csv('control_table.txt', index=None, sep=' ', mode='w')
		print('> finished')

		logging.info('Inventory rates calculated')

		#display results
		Label(window, text='Rates Calculated').grid(row=19,column=2)
		report_button = Button(window,
							text="Open Report in Excel",
							command = self.openReport)
		report_button.grid(row=20,column=2)
		num_prods = "Total Products: "+str(i.prod_table.shape[0])
		num_facs = "Total Facilities: "+str(i.fac_table.shape[0])
		Label(window, text=num_prods).grid(row=21,column=2)
		Label(window, text=num_facs).grid(row=22,column=2)
		config_button = Button(window,
							text="Save My Settings",
							command = self.createConfig)
		config_button.grid(row = 23,column=2)

	def openReport(self):
		
		os.system("start excel.exe ..\Inventory_Turn_Analysis(version2).xlsm")
		logging.info('Report created')

	def openSettings(self):
		# Settings panel where user can set directory for output & other settings
		window2=Tk()
		window2.title("Inventory Analysis Settings")

	def createConfig(self):
		config = configparser.ConfigParser()
		config.add_section('column headers')
		config.set('column headers','data_type', self.min_value.get())
		config.set('column headers','data_type', self.max_value.get())
		config.set('column headers','data_type', self.del_freq.get())
		config.set('column headers','data_type', self.data_type.get())
		config.set('column headers','product',self.product.get())
		config.set('column headers','product1',self.product1.get())
		config.set('column headers','product2',self.product2.get())
		config.set('column headers','facility',self.facility.get())
		config.set('column headers','facility1',self.facility1.get())
		config.set('column headers','facility2',self.facility2.get())
		config.set('column headers','date_field',self.date_field.get())
		config.set('column headers','issue_col',self.issue_col.get())
		config.set('column headers','soh_col',self.soh_col.get())
		# config.set('column headers','stockout_col',self.stockout_col.get())

		config.add_section('other')
		config.set('other','rates_window',str(self.rates_window.get()))
		config.set('other','issue_0',str(self.issue_0.get()))
		config.set('other','soh_0',str(self.soh_0.get()))
		# config.set('other','stockout_0',str(self.stockout_0.get()))

		basepath = path.abspath("")

		with open(path.join(basepath,'configfile.ini'),'w') as configfile:
			config.write(configfile)


IAGui(window)
window.mainloop()


