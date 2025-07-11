import tkinter
import customtkinter
from tkcalendar import Calendar

class CTkcalendar():
# start date title
self.start_date_title = customtkinter.CTkLabel(self.analysis_tab_view.tab('Overall Distribution'), text="Select Start Date", font=self.font_navigation_label)
self.start_date_title.grid(row=0, column=0, columnspan=1, padx=20, pady=10)

# start date calander
self.start_cal = Calendar(self.analysis_tab_view.tab('Overall Distribution'), selectmode='day', font=self.font_calender,
showweeknumbers=False, cursor="hand2", date_pattern= 'y-mm-dd',
borderwidth=0, bordercolor='white')
self.start_cal.grid(row= 1,column=0, padx=30, pady=10, sticky='w')

# end date title
self.end_date_title = customtkinter.CTkLabel(self.analysis_tab_view.tab('Overall Distribution'), text="Select End Date", font=self.font_navigation_label)
self.end_date_title.grid(row=0, column=1, columnspan=1, padx=20, pady=10)

# end date calander
self.end_cal = Calendar(self.analysis_tab_view.tab('Overall Distribution'), selectmode='day', font=self.font_calender,
showweeknumbers=False, cursor="hand2", date_pattern= 'y-mm-dd',
borderwidth=0, bordercolor='white')
self.end_cal.grid(row= 1,column=1, padx=30, pady=10, sticky='e')

# date confirm button
self.confirm_date = customtkinter.CTkButton(self.analysis_tab_view.tab('Overall Distribution'), text="Confirm Dates", font=self.font_details, hover=True, command=self.fetch_dates_overall)
self.confirm_date.grid(row=2, column=0, columnspan=2, padx=10, pady=30)

# start date entry
self.start_date_entry = customtkinter.CTkEntry(self.analysis_tab_view.tab('Overall Distribution'), placeholder_text="Start Date", font=self.font_details)
self.start_date_entry.grid(row=3, column=0, padx=10, pady=10)

# end date entry
self.end_date_entry = customtkinter.CTkEntry(self.analysis_tab_view.tab('Overall Distribution'), placeholder_text="End Date", font=self.font_details)
self.end_date_entry.grid(row=3, column=1, padx=10, pady=10)

self.proceed_overall_graph = customtkinter.CTkButton(self.analysis_tab_view.tab('Overall Distribution'))