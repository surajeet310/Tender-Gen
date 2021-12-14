import wx
import re
import os
import pickle
from threading import Thread
from TenderGen import TenderGen


class TenderGenApp(wx.Frame):
    def __init__(self, parent, title):
        super(TenderGenApp, self).__init__(
            parent,
            title=title,
            style=wx.SYSTEM_MENU | wx.CAPTION | wx.CLOSE_BOX | wx.MINIMIZE_BOX,
        )
        self.ID_review_gen = wx.NewIdRef()
        self.ID_doc_gen = wx.NewIdRef()
        self.ID_pdf_gen = wx.NewIdRef()
        self.ID_settings = wx.NewIdRef()
        self.ID_clear = wx.NewIdRef()
        self.status_bar = self.CreateStatusBar()
        self.UI()
        self.Center()

    def UI(self):
        panel = wx.Panel(self, style=wx.BORDER_THEME)
        panel.SetBackgroundColour((248, 249, 250))
        main_sizer = wx.GridBagSizer(5, 5)
        tender_section_sizer = wx.FlexGridSizer(9, 2, 15, 120)
        bidder_section_sizer = wx.FlexGridSizer(11, 2, 15, 65)
        others_section_sizer = wx.FlexGridSizer(3, 2, 15, 25)
        wx.Font.AddPrivateFont("fonts/Roboto-Regular.ttf")
        font = wx.Font(
            30, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD, False
        )

        # Menu Bar Init

        menubar = wx.MenuBar()
        file_menu_gen = wx.Menu()
        file_menu_clear = wx.Menu()
        file_menu_settings = wx.Menu()
        file_item_review_generate = file_menu_gen.Append(
            self.ID_review_gen,
            "Review and Generate",
            "Review details and generate file",
        )
        file_menu_gen.Bind(wx.EVT_ENTER_WINDOW, self.onWidgetEnter)
        file_item_generate_doc = file_menu_gen.Append(
            self.ID_doc_gen, "Generate Doc File", "Generate file in docx format"
        )
        file_item_generate_pdf = file_menu_gen.Append(
            self.ID_pdf_gen, "Generate PDF File", "Generate file in PDF format"
        )
        file_item_settings = file_menu_settings.Append(
            self.ID_settings, "Settings", "Settings"
        )
        file_item_clear_input = file_menu_clear.Append(
            self.ID_clear, "Clear All", "Clear all inputs"
        )
        menubar.Append(file_menu_gen, "&Action")
        menubar.Append(file_menu_settings, "&Settings")
        menubar.Append(file_menu_clear, "&Clear")
        self.SetMenuBar(menubar)
        self.Bind(wx.EVT_MENU, self.onMenuOptionSelected, id=self.ID_review_gen)
        self.Bind(wx.EVT_MENU, self.onMenuOptionSelected, id=self.ID_doc_gen)
        self.Bind(wx.EVT_MENU, self.onMenuOptionSelected, id=self.ID_pdf_gen)
        self.Bind(wx.EVT_MENU, self.onMenuOptionSelected, id=self.ID_settings)
        self.Bind(wx.EVT_MENU, self.onMenuOptionSelected, id=self.ID_clear)

        # Application Header

        app_text = wx.StaticText(panel, label="Tender Gen")
        app_text.SetOwnFont(font)
        main_sizer.Add(
            app_text,
            pos=(0, 0),
            flag=wx.EXPAND | wx.LEFT | wx.TOP | wx.BOTTOM,
            border=10,
        )
        app_icon = wx.StaticBitmap(panel, bitmap=wx.Bitmap("icons/docImg.png"))
        main_sizer.Add(
            app_icon,
            pos=(0, 1),
            flag=wx.TOP | wx.RIGHT | wx.ALIGN_RIGHT,
            border=10,
        )
        horizontal_line = wx.StaticLine(panel)
        main_sizer.Add(
            horizontal_line,
            pos=(1, 0),
            span=(1, 3),
            flag=wx.EXPAND | wx.BOTTOM,
            border=10,
        )
        # Tender Section

        tender_section_label = wx.StaticBox(panel, label="Tender Details")
        tender_details_sizer = wx.StaticBoxSizer(tender_section_label, wx.VERTICAL)

        tender_ref_num_text = wx.StaticText(panel, label="Reference Number")
        self.tender_ref_num = wx.TextCtrl(
            panel, size=(300, -1), name="Reference Number"
        )
        self.tender_ref_num.Bind(wx.EVT_ENTER_WINDOW, self.onWidgetEnter)
        self.tender_ref_num.Bind(wx.EVT_TEXT, self.tender_ref_num_reset)

        tender_name_of_work_text = wx.StaticText(panel, label="Name of Work")
        self.tender_name_of_work = wx.TextCtrl(panel, name="Name of Work")
        self.tender_name_of_work.Bind(wx.EVT_ENTER_WINDOW, self.onWidgetEnter)
        self.tender_name_of_work.Bind(wx.EVT_TEXT, self.tender_name_of_work_reset)

        tender_employer_name_text = wx.StaticText(panel, label="Employer's Name")
        self.tender_employer_name = wx.TextCtrl(panel, name="Employer's Name")
        self.tender_employer_name.Bind(wx.EVT_ENTER_WINDOW, self.onWidgetEnter)
        self.tender_employer_name.Bind(wx.EVT_TEXT, self.tender_emp_name_reset)

        tender_pckg_num_text = wx.StaticText(panel, label="Package Number")
        self.tender_pckg_num = wx.TextCtrl(panel, name="Package Number")
        self.tender_pckg_num.Bind(wx.EVT_ENTER_WINDOW, self.onWidgetEnter)
        self.tender_pckg_num.Bind(wx.EVT_TEXT, self.tender_pckg_num_reset)

        tender_estm_cost_text = wx.StaticText(panel, label="Estimated Cost")
        self.tender_estm_cost = wx.TextCtrl(panel, name="Estimated Cost")
        self.tender_estm_cost.SetHint("In Rupees")
        self.tender_estm_cost.Bind(wx.EVT_ENTER_WINDOW, self.onWidgetEnter)
        self.tender_estm_cost.Bind(wx.EVT_TEXT, self.tender_estm_cost_reset)

        tender_earnest_money_text = wx.StaticText(panel, label="Earnest Money Amount")
        self.tender_earnest_money = wx.TextCtrl(panel, name="Earnest Money Amount")
        self.tender_earnest_money.SetHint("In Rupees")
        self.tender_earnest_money.Bind(wx.EVT_ENTER_WINDOW, self.onWidgetEnter)
        self.tender_earnest_money.Bind(wx.EVT_TEXT, self.tender_ernst_money_reset)

        tender_paper_cost_text = wx.StaticText(panel, label="Tender Paper Cost")
        self.tender_paper_cost = wx.TextCtrl(panel, name="Tender Paper Cost")
        self.tender_paper_cost.SetHint("In Rupees")
        self.tender_paper_cost.Bind(wx.EVT_ENTER_WINDOW, self.onWidgetEnter)
        self.tender_paper_cost.Bind(wx.EVT_TEXT, self.tender_paper_cost_reset)

        tender_duration_text = wx.StaticText(panel, label="Duration of Work")
        tender_duration_box = wx.BoxSizer(wx.HORIZONTAL)
        self.tender_duration_years = wx.TextCtrl(panel, name="Duration of Work(Years)")
        self.tender_duration_years.SetHint("Year")

        self.tender_duration_months = wx.TextCtrl(
            panel, name="Duration of Work(Months)"
        )
        self.tender_duration_months.SetHint("Month")

        self.tender_duration_days = wx.TextCtrl(panel, name="Duration of Work(Days)")
        self.tender_duration_days.SetHint("Day")

        tender_duration_box.Add(self.tender_duration_years, flag=wx.ALL, border=4)
        tender_duration_box.Add(self.tender_duration_months, flag=wx.ALL, border=4)
        tender_duration_box.Add(self.tender_duration_days, flag=wx.ALL, border=4)

        self.tender_duration_years.Bind(wx.EVT_ENTER_WINDOW, self.onWidgetEnter)
        self.tender_duration_months.Bind(wx.EVT_ENTER_WINDOW, self.onWidgetEnter)
        self.tender_duration_days.Bind(wx.EVT_ENTER_WINDOW, self.onWidgetEnter)
        self.tender_duration_years.Bind(wx.EVT_TEXT, self.tender_duration_years_reset)
        self.tender_duration_months.Bind(wx.EVT_TEXT, self.tender_duration_months_reset)
        self.tender_duration_days.Bind(wx.EVT_TEXT, self.tender_duration_days_reset)

        bid_validity_text = wx.StaticText(panel, label="Bid Validity")
        bid_validity_box = wx.BoxSizer(wx.HORIZONTAL)
        self.bid_validity_years = wx.TextCtrl(panel, name="Bid Validity(Years)")
        self.bid_validity_years.SetHint("Year")

        self.bid_validity_months = wx.TextCtrl(panel, name="Bid Validity(Months)")
        self.bid_validity_months.SetHint("Month")

        self.bid_validity_days = wx.TextCtrl(panel, name="Bid Validity(Days)")
        self.bid_validity_days.SetHint("Day")

        bid_validity_box.Add(self.bid_validity_years, flag=wx.ALL, border=4)
        bid_validity_box.Add(self.bid_validity_months, flag=wx.ALL, border=4)
        bid_validity_box.Add(self.bid_validity_days, flag=wx.ALL, border=4)

        self.bid_validity_years.Bind(wx.EVT_ENTER_WINDOW, self.onWidgetEnter)
        self.bid_validity_months.Bind(wx.EVT_ENTER_WINDOW, self.onWidgetEnter)
        self.bid_validity_days.Bind(wx.EVT_ENTER_WINDOW, self.onWidgetEnter)
        self.bid_validity_years.Bind(wx.EVT_TEXT, self.bid_validity_years_reset)
        self.bid_validity_months.Bind(wx.EVT_TEXT, self.bid_validity_months_reset)
        self.bid_validity_days.Bind(wx.EVT_TEXT, self.bid_validity_days_reset)

        widget_list_tender = [
            (tender_ref_num_text),
            (self.tender_ref_num, 1, wx.EXPAND),
            (tender_name_of_work_text),
            (self.tender_name_of_work, 1, wx.EXPAND),
            (tender_employer_name_text),
            (self.tender_employer_name, 1, wx.EXPAND),
            (tender_pckg_num_text),
            (self.tender_pckg_num, 1, wx.EXPAND),
            (tender_estm_cost_text),
            (self.tender_estm_cost, 1, wx.EXPAND),
            (tender_earnest_money_text),
            (self.tender_earnest_money, 1, wx.EXPAND),
            (tender_paper_cost_text),
            (self.tender_paper_cost, 1, wx.EXPAND),
            (tender_duration_text),
            (tender_duration_box, 1, wx.EXPAND),
            (bid_validity_text),
            (bid_validity_box, 1, wx.EXPAND),
        ]
        tender_section_sizer.AddMany(widget_list_tender)

        tender_details_sizer.Add(
            tender_section_sizer,
            flag=wx.EXPAND | wx.ALL,
            border=20,
        )

        main_sizer.Add(
            tender_details_sizer,
            pos=(2, 0),
            flag=wx.LEFT | wx.BOTTOM | wx.EXPAND | wx.RIGHT,
            border=15,
        )

        # Bidder Section

        bidder_section_label = wx.StaticBox(panel, label="Bidder Details")
        bidder_details_sizer = wx.StaticBoxSizer(bidder_section_label, wx.VERTICAL)

        bidding_type_text = wx.StaticText(panel, label="Bidding Type")
        self.bidding_type = wx.ComboBox(
            panel,
            choices=[
                "Choose a bidding type",
                "Individual",
                "Proprietor",
                "Partnership",
            ],
            name="Type of bidding",
            style=wx.CB_READONLY,
        )
        self.bidding_type.SetValue("Choose a bidding type")
        self.bidding_type.Bind(wx.EVT_ENTER_WINDOW, self.onWidgetEnter)
        self.bidding_type.Bind(wx.EVT_COMBOBOX, self.onSelectChoice)

        contractor_name_text = wx.StaticText(panel, label="Name of Contractor")
        self.contractor_name = wx.TextCtrl(
            panel, size=(300, -1), name="Name of Contractor"
        )
        self.contractor_name.Bind(wx.EVT_ENTER_WINDOW, self.onWidgetEnter)
        self.contractor_name.Bind(wx.EVT_TEXT, self.contractor_name_reset)

        bidder_name_text = wx.StaticText(panel, label="Name of the Bidder")
        self.bidder_name = wx.TextCtrl(panel, name="Name of the Bidder")
        self.bidder_name.Bind(wx.EVT_ENTER_WINDOW, self.onWidgetEnter)
        self.bidder_name.Bind(wx.EVT_TEXT, self.bidder_name_reset)

        reg_num_text = wx.StaticText(panel, label="PWD (Roads) Registration Number")
        self.reg_num = wx.TextCtrl(panel, name="PWD (Roads) Registration Number")
        self.reg_num.Bind(wx.EVT_ENTER_WINDOW, self.onWidgetEnter)
        self.reg_num.Bind(wx.EVT_TEXT, self.reg_num_reset)

        bidder_address_text = wx.StaticText(panel, label="Full address")
        self.bidder_address = wx.TextCtrl(panel, name="Full address of Bidder")
        self.bidder_address.Bind(wx.EVT_ENTER_WINDOW, self.onWidgetEnter)
        self.bidder_address.Bind(wx.EVT_TEXT, self.bidder_address_reset)

        bidder_ph_num_text = wx.StaticText(panel, label="Mobile Number")
        self.bidder_ph_num = wx.TextCtrl(panel, name="Mobile Number of Bidder")
        self.bidder_ph_num.Bind(wx.EVT_ENTER_WINDOW, self.onWidgetEnter)
        self.bidder_ph_num.Bind(wx.EVT_TEXT, self.bidder_ph_num_reset)

        bidder_email_id_text = wx.StaticText(panel, label="Email id")
        self.bidder_email_id = wx.TextCtrl(panel, name="Email id of Bidder")
        self.bidder_email_id.Bind(wx.EVT_ENTER_WINDOW, self.onWidgetEnter)
        self.bidder_email_id.Bind(wx.EVT_TEXT, self.bidder_email_reset)

        bidder_ac_num_text = wx.StaticText(panel, label="Bank Account Number")
        self.bidder_ac_num = wx.TextCtrl(panel, name="Bank Account Number of Bidder")
        self.bidder_ac_num.Bind(wx.EVT_ENTER_WINDOW, self.onWidgetEnter)
        self.bidder_ac_num.Bind(wx.EVT_TEXT, self.bidder_ac_num_reset)

        bidder_bank_name_text = wx.StaticText(panel, label="Bank's name")
        self.bidder_bank_name = wx.TextCtrl(panel, name="Bank's name  of Bidder")
        self.bidder_bank_name.Bind(wx.EVT_ENTER_WINDOW, self.onWidgetEnter)
        self.bidder_bank_name.Bind(wx.EVT_TEXT, self.bidder_bank_name_reset)

        bidder_bank_ifsc_text = wx.StaticText(panel, label="Bank IFSC code ")
        self.bidder_bank_ifsc = wx.TextCtrl(panel, name="Bank IFSC code for bidder")
        self.bidder_bank_ifsc.Bind(wx.EVT_ENTER_WINDOW, self.onWidgetEnter)
        self.bidder_bank_ifsc.Bind(wx.EVT_TEXT, self.bidder_bank_ifsc_reset)

        bidder_bank_branch_text = wx.StaticText(panel, label="Bank's Branch")
        self.bidder_bank_branch = wx.TextCtrl(panel, name="Bank's Branch  of Bidder")
        self.bidder_bank_branch.Bind(wx.EVT_ENTER_WINDOW, self.onWidgetEnter)
        self.bidder_bank_branch.Bind(wx.EVT_TEXT, self.bidder_bank_branch_reset)

        widget_list_bidder = [
            (contractor_name_text),
            (self.contractor_name, 1, wx.EXPAND),
            (bidding_type_text),
            (self.bidding_type, 1, wx.EXPAND),
            (bidder_name_text),
            (self.bidder_name, 1, wx.EXPAND),
            (reg_num_text),
            (self.reg_num, 1, wx.EXPAND),
            (bidder_address_text),
            (self.bidder_address, 1, wx.EXPAND),
            (bidder_ph_num_text),
            (self.bidder_ph_num, 1, wx.EXPAND),
            (bidder_email_id_text),
            (self.bidder_email_id, 1, wx.EXPAND),
            (bidder_ac_num_text),
            (self.bidder_ac_num, 1, wx.EXPAND),
            (bidder_bank_name_text),
            (self.bidder_bank_name, 1, wx.EXPAND),
            (bidder_bank_ifsc_text),
            (self.bidder_bank_ifsc, 1, wx.EXPAND),
            (bidder_bank_branch_text),
            (self.bidder_bank_branch, 1, wx.EXPAND),
        ]

        bidder_section_sizer.AddMany(widget_list_bidder)
        bidder_details_sizer.Add(
            bidder_section_sizer, flag=wx.EXPAND | wx.ALL, border=20
        )
        main_sizer.Add(
            bidder_details_sizer,
            pos=(2, 1),
            flag=wx.BOTTOM | wx.LEFT | wx.RIGHT | wx.EXPAND | wx.ALIGN_RIGHT,
            border=15,
        )

        # Other Details Section

        other_details_label = wx.StaticBox(panel, label="Other Details")
        other_details_sizer = wx.StaticBoxSizer(other_details_label, wx.VERTICAL)

        place_of_reg_text = wx.StaticText(panel, label="Place of registration")
        self.place_of_reg = wx.TextCtrl(
            panel, size=(250, -1), name="Place of registration"
        )
        self.place_of_reg.Bind(wx.EVT_ENTER_WINDOW, self.onWidgetEnter)
        self.place_of_reg.Bind(wx.EVT_TEXT, self.place_of_reg_reset)

        place_of_buisness_text = wx.StaticText(
            panel, label="Principal place of buisness"
        )
        self.place_of_buisness = wx.TextCtrl(panel, name="Principal place of buisness")
        self.place_of_buisness.Bind(wx.EVT_ENTER_WINDOW, self.onWidgetEnter)
        self.place_of_buisness.Bind(wx.EVT_TEXT, self.place_of_buisness_reset)

        cash_invest_text = wx.StaticText(
            panel, label="Min. cash investment(during implementation of contract)"
        )
        self.cash_invest = wx.TextCtrl(
            panel, name="Min. cash investment(during implementation of contract)"
        )
        self.cash_invest.SetHint("in %")
        self.cash_invest.Bind(wx.EVT_ENTER_WINDOW, self.onWidgetEnter)
        self.cash_invest.Bind(wx.EVT_TEXT, self.min_cash_reset)

        widget_list_others = [
            (place_of_reg_text),
            (self.place_of_reg, 1, wx.EXPAND),
            (place_of_buisness_text),
            (self.place_of_buisness, 1, wx.EXPAND),
            (cash_invest_text),
            (self.cash_invest, 1, wx.EXPAND),
        ]
        others_section_sizer.AddMany(widget_list_others)

        other_details_sizer.Add(
            others_section_sizer,
            flag=wx.EXPAND | wx.ALL,
            border=15,
        )

        main_sizer.Add(
            other_details_sizer,
            pos=(3, 0),
            flag=wx.LEFT | wx.BOTTOM | wx.EXPAND,
            border=15,
        )

        # Application Footer

        footer_line = wx.StaticLine(panel)
        main_sizer.Add(
            footer_line,
            pos=(4, 0),
            span=(4, 3),
            flag=wx.TOP | wx.BOTTOM | wx.EXPAND,
            border=20,
        )

        panel.SetSizer(main_sizer)
        main_sizer.Fit(self)

        self.Bind(wx.EVT_CLOSE, self.onCloseWindow)

    def onWidgetEnter(self, e):
        name = e.GetEventObject().GetName()
        self.status_bar.SetStatusText(name)
        e.Skip()

    def onMenuOptionSelected(self, e):
        option_id = e.GetId()
        if option_id == self.ID_settings:
            setting_obj = Settings(self)
            setting_obj.Show()
        if option_id == self.ID_clear:
            self.clearInputFields()
        if option_id == self.ID_review_gen:
            response = False
            validate_obj = ValidateData(self)
            if validate_obj.checkIfEmpty():
                response = (validate_obj.checkIfNumber()) & (
                    validate_obj.checkEmailId()
                )
            if response:
                review_window = ReviewFrame(self)
                review_window.Show()
        if option_id == self.ID_doc_gen:
            response = False
            validate_obj = ValidateData(self)
            if validate_obj.checkIfEmpty():
                response = (validate_obj.checkIfNumber()) & (
                    validate_obj.checkEmailId()
                )
            if response:
                gen_window_1 = GenerateFrame(self, False)
        if option_id == self.ID_pdf_gen:
            response = False
            validate_obj = ValidateData(self)
            if validate_obj.checkIfEmpty():
                response = (validate_obj.checkIfNumber()) & (
                    validate_obj.checkEmailId()
                )
            if response:
                gen_window_2 = GenerateFrame(self, True)

    def onSelectChoice(self, e):
        self.bidding_type.SetBackgroundColour((255, 255, 255))
        bidding_type_choice = e.GetString()
        if bidding_type_choice == "Individual":
            self.bidder_name.SetValue(self.contractor_name.GetValue())

    def onCloseWindow(self, e):
        dial = wx.MessageDialog(
            None,
            "Are you sure you want to quit ?",
            "Question",
            wx.YES_NO | wx.NO_DEFAULT | wx.ICON_QUESTION,
        )
        choice = dial.ShowModal()
        if choice == wx.ID_YES:
            self.Destroy()
        else:
            e.Veto()

    def clearInputFields(self):
        self.tender_ref_num.Clear()
        self.tender_name_of_work.Clear()
        self.tender_pckg_num.Clear()
        self.tender_employer_name.Clear()
        self.tender_estm_cost.Clear()
        self.tender_earnest_money.Clear()
        self.tender_paper_cost.Clear()
        self.tender_duration_years.Clear()
        self.tender_duration_months.Clear()
        self.tender_duration_days.Clear()
        self.bid_validity_years.Clear()
        self.bid_validity_months.Clear()
        self.bid_validity_days.Clear()
        self.contractor_name.Clear()
        self.bidding_type.SetValue("Choose a bidding type")
        self.bidder_name.Clear()
        self.reg_num.Clear()
        self.bidder_address.Clear()
        self.bidder_ph_num.Clear()
        self.bidder_email_id.Clear()
        self.bidder_ac_num.Clear()
        self.bidder_bank_name.Clear()
        self.bidder_bank_branch.Clear()
        self.bidder_bank_ifsc.Clear()
        self.place_of_reg.Clear()
        self.place_of_buisness.Clear()
        self.cash_invest.Clear()

    def tender_ref_num_reset(self, e):
        self.tender_ref_num.SetBackgroundColour((255, 255, 255))

    def tender_name_of_work_reset(self, e):
        self.tender_name_of_work.SetBackgroundColour((255, 255, 255))

    def tender_pckg_num_reset(self, e):
        self.tender_pckg_num.SetBackgroundColour((255, 255, 255))

    def tender_emp_name_reset(self, e):
        self.tender_employer_name.SetBackgroundColour((255, 255, 255))

    def tender_estm_cost_reset(self, e):
        self.tender_estm_cost.SetBackgroundColour((255, 255, 255))

    def tender_ernst_money_reset(self, e):
        self.tender_earnest_money.SetBackgroundColour((255, 255, 255))

    def tender_paper_cost_reset(self, e):
        self.tender_paper_cost.SetBackgroundColour((255, 255, 255))

    def tender_duration_years_reset(self, e):
        self.tender_duration_years.SetBackgroundColour((255, 255, 255))

    def tender_duration_months_reset(self, e):
        self.tender_duration_months.SetBackgroundColour((255, 255, 255))

    def tender_duration_days_reset(self, e):
        self.tender_duration_days.SetBackgroundColour((255, 255, 255))

    def bid_validity_years_reset(self, e):
        self.bid_validity_years.SetBackgroundColour((255, 255, 255))

    def bid_validity_months_reset(self, e):
        self.bid_validity_months.SetBackgroundColour((255, 255, 255))

    def bid_validity_days_reset(self, e):
        self.bid_validity_days.SetBackgroundColour((255, 255, 255))

    def contractor_name_reset(self, e):
        self.contractor_name.SetBackgroundColour((255, 255, 255))

    def bidder_name_reset(self, e):
        self.bidder_name.SetBackgroundColour((255, 255, 255))

    def reg_num_reset(self, e):
        self.reg_num.SetBackgroundColour((255, 255, 255))

    def bidder_address_reset(self, e):
        self.bidder_address.SetBackgroundColour((255, 255, 255))

    def bidder_ph_num_reset(self, e):
        self.bidder_ph_num.SetBackgroundColour((255, 255, 255))

    def bidder_email_reset(self, e):
        self.bidder_email_id.SetBackgroundColour((255, 255, 255))

    def bidder_ac_num_reset(self, e):
        self.bidder_ac_num.SetBackgroundColour((255, 255, 255))

    def bidder_bank_name_reset(self, e):
        self.bidder_bank_name.SetBackgroundColour((255, 255, 255))

    def bidder_bank_branch_reset(self, e):
        self.bidder_bank_branch.SetBackgroundColour((255, 255, 255))

    def bidder_bank_ifsc_reset(self, e):
        self.bidder_bank_ifsc.SetBackgroundColour((255, 255, 255))

    def place_of_reg_reset(self, e):
        self.place_of_reg.SetBackgroundColour((255, 255, 255))

    def place_of_buisness_reset(self, e):
        self.place_of_buisness.SetBackgroundColour((255, 255, 255))

    def min_cash_reset(self, e):
        self.cash_invest.SetBackgroundColour((255, 255, 255))


class ValidateData:
    def __init__(self, UI_object):
        self.UI = UI_object

    def checkIfEmpty(self):
        response = True
        if self.UI.tender_ref_num.GetValue() == "":
            self.UI.tender_ref_num.SetBackgroundColour((255, 51, 0))
            self.UI.tender_ref_num.SetFocus()
            self.UI.tender_ref_num.Refresh()
            response = False
        if self.UI.tender_name_of_work.GetValue() == "":
            self.UI.tender_name_of_work.SetBackgroundColour((255, 51, 0))
            self.UI.tender_name_of_work.SetFocus()
            self.UI.tender_name_of_work.Refresh()
            response = False
        if self.UI.tender_pckg_num.GetValue() == "":
            self.UI.tender_pckg_num.SetBackgroundColour((255, 51, 0))
            self.UI.tender_pckg_num.SetFocus()
            self.UI.tender_pckg_num.Refresh()
            response = False
        if self.UI.tender_employer_name.GetValue() == "":
            self.UI.tender_employer_name.SetBackgroundColour((255, 51, 0))
            self.UI.tender_employer_name.SetFocus()
            self.UI.tender_employer_name.Refresh()
            response = False
        if self.UI.tender_estm_cost.GetValue() == "":
            self.UI.tender_estm_cost.SetBackgroundColour((255, 51, 0))
            self.UI.tender_estm_cost.SetFocus()
            self.UI.tender_estm_cost.Refresh()
            response = False
        if self.UI.tender_earnest_money.GetValue() == "":
            self.UI.tender_earnest_money.SetBackgroundColour((255, 51, 0))
            self.UI.tender_earnest_money.SetFocus()
            self.UI.tender_earnest_money.Refresh()
            response = False
        if self.UI.tender_paper_cost.GetValue() == "":
            self.UI.tender_paper_cost.SetBackgroundColour((255, 51, 0))
            self.UI.tender_paper_cost.SetFocus()
            self.UI.tender_paper_cost.Refresh()
            response = False
        if self.UI.tender_duration_years.GetValue() == "":
            self.UI.tender_duration_years.SetBackgroundColour((255, 51, 0))
            self.UI.tender_duration_years.SetFocus()
            self.UI.tender_duration_years.Refresh()
            response = False
        if self.UI.tender_duration_months.GetValue() == "":
            self.UI.tender_duration_months.SetBackgroundColour((255, 51, 0))
            self.UI.tender_duration_months.SetFocus()
            self.UI.tender_duration_months.Refresh()
            response = False
        if self.UI.tender_duration_days.GetValue() == "":
            self.UI.tender_duration_days.SetBackgroundColour((255, 51, 0))
            self.UI.tender_duration_days.SetFocus()
            self.UI.tender_duration_days.Refresh()
            response = False
        if self.UI.bid_validity_years.GetValue() == "":
            self.UI.bid_validity_years.SetBackgroundColour((255, 51, 0))
            self.UI.bid_validity_years.SetFocus()
            self.UI.bid_validity_years.Refresh()
            response = False
        if self.UI.bid_validity_months.GetValue() == "":
            self.UI.bid_validity_months.SetBackgroundColour((255, 51, 0))
            self.UI.bid_validity_months.SetFocus()
            self.UI.bid_validity_months.Refresh()
            response = False
        if self.UI.bid_validity_days.GetValue() == "":
            self.UI.bid_validity_days.SetBackgroundColour((255, 51, 0))
            self.UI.bid_validity_days.SetFocus()
            self.UI.bid_validity_days.Refresh()
            response = False
        if self.UI.contractor_name.GetValue() == "":
            self.UI.contractor_name.SetBackgroundColour((255, 51, 0))
            self.UI.contractor_name.SetFocus()
            self.UI.contractor_name.Refresh()
            response = False
        if self.UI.bidding_type.GetValue() == "Choose a bidding type":
            self.UI.bidding_type.SetBackgroundColour((255, 51, 0))
            self.UI.bidding_type.SetFocus()
            self.UI.bidding_type.Refresh()
            response = False
        if self.UI.bidder_name.GetValue() == "":
            self.UI.bidder_name.SetBackgroundColour((255, 51, 0))
            self.UI.bidder_name.SetFocus()
            self.UI.bidder_name.Refresh()
            response = False
        if self.UI.reg_num.GetValue() == "":
            self.UI.reg_num.SetBackgroundColour((255, 51, 0))
            self.UI.reg_num.SetFocus()
            self.UI.reg_num.Refresh()
            response = False
        if self.UI.bidder_address.GetValue() == "":
            self.UI.bidder_address.SetBackgroundColour((255, 51, 0))
            self.UI.bidder_address.SetFocus()
            self.UI.bidder_address.Refresh()
            response = False
        if self.UI.bidder_ph_num.GetValue() == "":
            self.UI.bidder_ph_num.SetBackgroundColour((255, 51, 0))
            self.UI.bidder_ph_num.SetFocus()
            self.UI.bidder_ph_num.Refresh()
            response = False
        if self.UI.bidder_email_id.GetValue() == "":
            self.UI.bidder_email_id.SetBackgroundColour((255, 51, 0))
            self.UI.bidder_email_id.SetFocus()
            self.UI.bidder_email_id.Refresh()
            response = False
        if self.UI.bidder_ac_num.GetValue() == "":
            self.UI.bidder_ac_num.SetBackgroundColour((255, 51, 0))
            self.UI.bidder_ac_num.SetFocus()
            self.UI.bidder_ac_num.Refresh()
            response = False
        if self.UI.bidder_bank_name.GetValue() == "":
            self.UI.bidder_bank_name.SetBackgroundColour((255, 51, 0))
            self.UI.bidder_bank_name.SetFocus()
            self.UI.bidder_bank_name.Refresh()
            response = False
        if self.UI.bidder_bank_branch.GetValue() == "":
            self.UI.bidder_bank_branch.SetBackgroundColour((255, 51, 0))
            self.UI.bidder_bank_branch.SetFocus()
            self.UI.bidder_bank_branch.Refresh()
            response = False
        if self.UI.bidder_bank_ifsc.GetValue() == "":
            self.UI.bidder_bank_ifsc.SetBackgroundColour((255, 51, 0))
            self.UI.bidder_bank_ifsc.SetFocus()
            self.UI.bidder_bank_ifsc.Refresh()
            response = False
        if self.UI.place_of_reg.GetValue() == "":
            self.UI.place_of_reg.SetBackgroundColour((255, 51, 0))
            self.UI.place_of_reg.SetFocus()
            self.UI.place_of_reg.Refresh()
            response = False
        if self.UI.place_of_buisness.GetValue() == "":
            self.UI.place_of_buisness.SetBackgroundColour((255, 51, 0))
            self.UI.place_of_buisness.SetFocus()
            self.UI.place_of_buisness.Refresh()
            response = False
        if self.UI.cash_invest.GetValue() == "":
            self.UI.cash_invest.SetBackgroundColour((255, 51, 0))
            self.UI.cash_invest.SetFocus()
            self.UI.cash_invest.Refresh()
            response = False

        if response is not True:
            self.UI.status_bar.SetStatusText("Fill all the text fields.")
            return False
        else:
            return True

    def checkIfNumber(self):
        response = True
        try:
            int(self.UI.tender_estm_cost.GetValue())
        except ValueError:
            self.UI.tender_estm_cost.SetBackgroundColour((255, 51, 0))
            self.UI.tender_estm_cost.SetFocus()
            self.UI.tender_estm_cost.Refresh()
            response = False
        try:
            int(self.UI.tender_earnest_money.GetValue())
        except ValueError:
            self.UI.tender_earnest_money.SetBackgroundColour((255, 51, 0))
            self.UI.tender_earnest_money.SetFocus()
            self.UI.tender_earnest_money.Refresh()
            response = False
        try:
            int(self.UI.tender_paper_cost.GetValue())
        except ValueError:
            self.UI.tender_paper_cost.SetBackgroundColour((255, 51, 0))
            self.UI.tender_paper_cost.SetFocus()
            self.UI.tender_paper_cost.Refresh()
            response = False
        try:
            int(self.UI.tender_duration_years.GetValue())
        except ValueError:
            self.UI.tender_duration_years.SetBackgroundColour((255, 51, 0))
            self.UI.tender_duration_years.SetFocus()
            self.UI.tender_duration_years.Refresh()
            response = False
        try:
            int(self.UI.tender_duration_months.GetValue())
        except ValueError:
            self.UI.tender_duration_months.SetBackgroundColour((255, 51, 0))
            self.UI.tender_duration_months.SetFocus()
            self.UI.tender_duration_months.Refresh()
            response = False
        try:
            int(self.UI.tender_duration_days.GetValue())
        except ValueError:
            self.UI.tender_duration_days.SetBackgroundColour((255, 51, 0))
            self.UI.tender_duration_days.SetFocus()
            self.UI.tender_duration_days.Refresh()
            response = False
        try:
            int(self.UI.bid_validity_years.GetValue())
        except ValueError:
            self.UI.bid_validity_years.SetBackgroundColour((255, 51, 0))
            self.UI.bid_validity_years.SetFocus()
            self.UI.bid_validity_years.Refresh()
            response = False
        try:
            int(self.UI.bid_validity_months.GetValue())
        except ValueError:
            self.UI.bid_validity_months.SetBackgroundColour((255, 51, 0))
            self.UI.bid_validity_months.SetFocus()
            self.UI.bid_validity_months.Refresh()
            response = False
        try:
            int(self.UI.bid_validity_days.GetValue())
        except ValueError:
            self.UI.bid_validity_days.SetBackgroundColour((255, 51, 0))
            self.UI.bid_validity_days.SetFocus()
            self.UI.bid_validity_days.Refresh()
            response = False
        try:
            int(self.UI.bidder_ph_num.GetValue())
            if len(self.UI.bidder_ph_num.GetValue()) != 10:
                self.UI.bidder_ph_num.SetBackgroundColour((255, 51, 0))
                self.UI.bidder_ph_num.SetFocus()
                self.UI.bidder_ph_num.Refresh()
                response = False
        except ValueError:
            self.UI.bidder_ph_num.SetBackgroundColour((255, 51, 0))
            self.UI.bidder_ph_num.SetFocus()
            self.UI.bidder_ph_num.Refresh()
            response = False
        try:
            int(self.UI.bidder_ac_num.GetValue())
        except ValueError:
            self.UI.bidder_ac_num.SetBackgroundColour((255, 51, 0))
            self.UI.bidder_ac_num.SetFocus()
            self.UI.bidder_ac_num.Refresh()
            response = False
        try:
            int(self.UI.cash_invest.GetValue())
            if (int(self.UI.cash_invest.GetValue())) > 100:
                self.UI.cash_invest.SetBackgroundColour((255, 51, 0))
                self.UI.cash_invest.SetFocus()
                self.UI.cash_invest.Refresh()
                response = False
        except ValueError:
            self.UI.cash_invest.SetBackgroundColour((255, 51, 0))
            self.UI.cash_invest.SetFocus()
            self.UI.cash_invest.Refresh()
            response = False

        if response is not True:
            self.UI.status_bar.SetStatusText("Invalid details entered.")
            return False
        else:
            return True

    def checkEmailId(self):
        email = self.UI.bidder_email_id.GetValue()
        regex = r"\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b"
        if re.fullmatch(regex, email):
            return True
        else:
            self.UI.bidder_email_id.SetBackgroundColour((255, 51, 0))
            self.UI.bidder_email_id.SetFocus()
            self.UI.bidder_email_id.Refresh()
            self.UI.status_bar.SetStatusText("Invalid email id entered.")
            return False


class ReviewFrame(wx.Frame):
    def __init__(self, main_window):
        self.parent = main_window
        super(ReviewFrame, self).__init__(
            self.parent,
            title="Review Data",
            style=wx.SYSTEM_MENU | wx.CAPTION | wx.CLOSE_BOX,
        )
        self.InitUI()
        self.Center()

    def InitUI(self):
        panel = wx.Panel(self)
        global_sizer = wx.GridBagSizer(5, 5)
        tender_section_sizer = wx.FlexGridSizer(9, 2, 15, 20)
        bidder_section_sizer = wx.FlexGridSizer(14, 2, 15, 20)
        buttons_sizer = wx.BoxSizer(wx.HORIZONTAL)
        font = wx.Font(
            13, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD, False
        )

        tender_label = wx.StaticText(panel, label="Tender Details")
        bidder_label = wx.StaticText(panel, label="Bidder Details")
        tender_label.SetOwnFont(font)
        bidder_label.SetOwnFont(font)
        header_line = wx.StaticLine(panel)
        footer_line = wx.StaticLine(panel)
        global_sizer.Add(
            tender_label,
            pos=(0, 0),
            flag=wx.EXPAND | wx.TOP | wx.LEFT,
            border=15,
        )
        global_sizer.Add(
            bidder_label,
            pos=(0, 1),
            flag=wx.EXPAND | wx.TOP | wx.LEFT,
            border=15,
        )
        global_sizer.Add(
            header_line, pos=(1, 0), span=(1, 3), flag=wx.TOP | wx.EXPAND, border=15
        )

        # Tender Details Review

        tender_ref_num = wx.StaticText(panel, label="Tender reference Number :")
        tender_ref_num_res = wx.StaticText(
            panel, label=self.parent.tender_ref_num.GetValue()
        )

        tender_name_of_work = wx.StaticText(panel, label="Name of the work :")
        tender_name_of_work_res = wx.StaticText(
            panel, label=self.parent.tender_name_of_work.GetValue()
        )
        tender_name_of_work_res.Wrap(300)

        tender_pckg_num = wx.StaticText(panel, label="Package Number :")
        tender_pckg_num_res = wx.StaticText(
            panel, label=self.parent.tender_pckg_num.GetValue()
        )

        tender_emp_name = wx.StaticText(panel, label="Employer Name :")
        tender_emp_name_res = wx.StaticText(
            panel, label=self.parent.tender_employer_name.GetValue()
        )
        tender_emp_name_res.Wrap(300)

        tender_estm_cost = wx.StaticText(panel, label="Tender estimated cost :")
        tender_estm_cost_res = wx.StaticText(
            panel, label=self.parent.tender_estm_cost.GetValue()
        )

        tender_ernst_money = wx.StaticText(panel, label="Tender Earnest Money :")
        tender_ernst_money_res = wx.StaticText(
            panel, label=self.parent.tender_earnest_money.GetValue()
        )

        tender_paper_cost = wx.StaticText(panel, label="Tender paper cost :")
        tender_paper_cost_res = wx.StaticText(
            panel, label=self.parent.tender_paper_cost.GetValue()
        )

        tender_duration = wx.StaticText(panel, label="Duration of completion:")

        duration = list()
        if int(self.parent.tender_duration_years.GetValue()) != 0:
            duration.append(self.parent.tender_duration_years.GetValue() + " Years")
        if int(self.parent.tender_duration_months.GetValue()) != 0:
            duration.append(self.parent.tender_duration_months.GetValue() + " Months")
        if int(self.parent.tender_duration_days.GetValue()) != 0:
            duration.append(self.parent.tender_duration_days.GetValue() + " Days")

        tender_duration_res = wx.StaticText(panel, label=" ".join(duration))

        bid_validity = wx.StaticText(panel, label="Bid validity:")

        validity = list()
        if int(self.parent.bid_validity_years.GetValue()) != 0:
            validity.append(self.parent.bid_validity_years.GetValue() + " Years")
        if int(self.parent.bid_validity_months.GetValue()) != 0:
            validity.append(self.parent.bid_validity_months.GetValue() + " Months")
        if int(self.parent.bid_validity_days.GetValue()) != 0:
            validity.append(self.parent.bid_validity_days.GetValue() + " Days")

        bid_validity_res = wx.StaticText(panel, label=" ".join(validity))

        widgets = [
            (tender_ref_num),
            (tender_ref_num_res),
            (tender_name_of_work),
            (tender_name_of_work_res),
            (tender_pckg_num),
            (tender_pckg_num_res),
            (tender_emp_name),
            (tender_emp_name_res),
            (tender_estm_cost),
            (tender_estm_cost_res),
            (tender_ernst_money),
            (tender_ernst_money_res),
            (tender_paper_cost),
            (tender_paper_cost_res),
            (tender_duration),
            (tender_duration_res),
            (bid_validity),
            (bid_validity_res),
        ]
        tender_section_sizer.AddMany(widgets)
        global_sizer.Add(
            tender_section_sizer, pos=(2, 0), flag=wx.ALL | wx.EXPAND, border=25
        )

        # Bidding details review

        bidding_type = wx.StaticText(panel, label="Bidding Type")
        bidding_type_res = wx.StaticText(
            panel, label=self.parent.bidding_type.GetValue()
        )

        contractor_name = wx.StaticText(panel, label="Contractor Name :")
        contractor_name_res = wx.StaticText(
            panel, label=self.parent.contractor_name.GetValue()
        )

        bidder_name = wx.StaticText(panel, label="Bidder Name :")
        bidder_name_res = wx.StaticText(panel, label=self.parent.bidder_name.GetValue())

        reg_num = wx.StaticText(panel, label="PWD (Roads) Registration Number :")
        reg_num_res = wx.StaticText(panel, label=self.parent.reg_num.GetValue())

        bidder_address = wx.StaticText(panel, label="Full Address of Bidder :")
        bidder_address_res = wx.StaticText(
            panel, label=self.parent.bidder_address.GetValue()
        )
        bidder_address_res.Wrap(300)

        mob_num = wx.StaticText(panel, label="Bidder's Mobile Number :")
        mob_num_res = wx.StaticText(panel, label=self.parent.bidder_ph_num.GetValue())

        email = wx.StaticText(panel, label="Bidder's Email id :")
        email_res = wx.StaticText(panel, label=self.parent.bidder_email_id.GetValue())

        bidder_bank_ac = wx.StaticText(panel, label="Bank Account Number :")
        bidder_bank_ac_res = wx.StaticText(
            panel, label=self.parent.bidder_ac_num.GetValue()
        )

        bidder_bank_name = wx.StaticText(panel, label="Name of the Bank :")
        bidder_bank_name_res = wx.StaticText(
            panel, label=self.parent.bidder_bank_name.GetValue()
        )

        bidder_bank_branch = wx.StaticText(panel, label="Bank Branch :")
        bidder_bank_branch_res = wx.StaticText(
            panel, label=self.parent.bidder_bank_branch.GetValue()
        )

        bidder_bank_ifsc = wx.StaticText(panel, label="Bank IFSC Code :")
        bidder_bank_ifsc_res = wx.StaticText(
            panel, label=self.parent.bidder_bank_ifsc.GetValue()
        )

        place_reg = wx.StaticText(panel, label="Place of registration :")
        place_reg_res = wx.StaticText(panel, label=self.parent.place_of_reg.GetValue())
        place_reg_res.Wrap(300)

        place_buisness = wx.StaticText(panel, label="Place of buisness :")
        place_buisness_res = wx.StaticText(
            panel, label=self.parent.place_of_buisness.GetValue()
        )

        cash_invest = wx.StaticText(
            panel,
            label="Min. cash investment during implementation of contract (%) :",
        )
        cash_invest_res = wx.StaticText(panel, label=self.parent.cash_invest.GetValue())

        widget_bidder = [
            (bidding_type),
            (bidding_type_res),
            (contractor_name),
            (contractor_name_res),
            (bidder_name),
            (bidder_name_res),
            (reg_num),
            (reg_num_res),
            (bidder_address),
            (bidder_address_res),
            (mob_num),
            (mob_num_res),
            (email),
            (email_res),
            (bidder_bank_ac),
            (bidder_bank_ac_res),
            (bidder_bank_ifsc),
            (bidder_bank_ifsc_res),
            (bidder_bank_name),
            (bidder_bank_name_res),
            (bidder_bank_branch),
            (bidder_bank_branch_res),
            (place_reg),
            (place_reg_res),
            (place_buisness),
            (place_buisness_res),
            (cash_invest),
            (cash_invest_res),
        ]
        bidder_section_sizer.AddMany(widget_bidder)
        global_sizer.Add(
            bidder_section_sizer, pos=(2, 1), flag=wx.ALL | wx.EXPAND, border=25
        )

        pdf_gen_btn = wx.Button(panel, label="Generate PDF file")
        doc_gen_btn = wx.Button(panel, label="Generate Doc file")
        doc_gen_btn.Bind(wx.EVT_BUTTON, self.onClickDoc)
        pdf_gen_btn.Bind(wx.EVT_BUTTON, self.onClickPdf)
        buttons_sizer.Add(pdf_gen_btn, flag=wx.ALL, border=20)
        buttons_sizer.Add(doc_gen_btn, flag=wx.ALL, border=20)

        global_sizer.Add(
            buttons_sizer, pos=(3, 1), flag=wx.ALL | wx.ALIGN_CENTER, border=20
        )

        global_sizer.Add(
            footer_line, pos=(4, 0), span=(4, 3), flag=wx.TOP | wx.EXPAND, border=20
        )

        panel.SetSizer(global_sizer)
        global_sizer.Fit(self)

    def onClickDoc(self, e):
        gen_window = GenerateFrame(self.parent, False)
        self.Destroy()

    def onClickPdf(self, e):
        gen_window = GenerateFrame(self.parent, True)
        self.Destroy()


class GenerateFrame(wx.Frame):
    def __init__(self, main_window, pdf):
        self.parent_frame = main_window
        self.pdf = pdf
        super(GenerateFrame, self).__init__(
            self.parent_frame,
            title="Generate File",
            style=wx.SYSTEM_MENU | wx.CAPTION,
        )
        self.Center()
        self.InitUI()

    def InitUI(self):
        self.timer = wx.Timer(self, 1)
        self.count = 0
        self.task_range = 100
        self.sizer = wx.BoxSizer(wx.VERTICAL)
        self.panel = wx.Panel(self)
        self.gauge = wx.Gauge(self.panel, range=self.task_range, size=(300, -1))
        self.text_msg = wx.StaticText(self.panel, label="Processing...")
        self.Bind(wx.EVT_TIMER, self.onTimerUpdate, self.timer)
        self.ok_btn = wx.Button(self.panel, label="Ok")
        self.ok_btn.Bind(wx.EVT_BUTTON, self.onClose)
        self.ok_btn.Disable()
        self.success_icon = wx.StaticBitmap(
            self.panel, bitmap=wx.Bitmap("icons/success.png")
        )
        self.success_icon.Hide()
        self.fail_icon = wx.StaticBitmap(self.panel, bitmap=wx.Bitmap("icons/fail.png"))
        self.fail_icon.Hide()
        self.sizer.Add(self.gauge, flag=wx.ALL | wx.EXPAND, border=50)
        self.sizer.Add(self.text_msg, flag=wx.ALL | wx.ALIGN_CENTER, border=50)
        self.sizer.Add(self.success_icon, flag=wx.ALL | wx.ALIGN_CENTER, border=20)
        self.sizer.Add(self.fail_icon, flag=wx.ALL | wx.ALIGN_CENTER, border=20)
        self.sizer.Add(self.ok_btn, flag=wx.ALL | wx.ALIGN_CENTER, border=20)

        self.timer.Start(100)
        # thread 1 (doc)
        self.gen_obj = TenderGen(self.parent_frame, self.pdf)
        t1 = Thread(target=self.gen_obj.generate)
        t1.start()

        self.panel.SetSizer(self.sizer)
        self.sizer.Fit(self)
        self.Show()

    def onTimerUpdate(self, e):
        self.count += 10
        self.gauge.SetValue(self.count)
        if self.gen_obj.status is not None:
            self.timer.Stop()
            self.gauge.Hide()
            if self.gen_obj.status is True:
                self.text_msg.SetLabel("File Generation Successful")
                self.success_icon.Show()
                self.sizer.Layout()
            else:
                self.text_msg.SetLabel("File Generation Failed")
                self.fail_icon.Show()
                self.sizer.Layout()
            self.ok_btn.Enable()

        if self.count == self.task_range:
            self.count = 0

    def onClose(self, e):
        self.Destroy()


class Settings(wx.Frame):
    def __init__(self, parent_win):
        self.parent_win = parent_win
        super(Settings, self).__init__(
            self.parent_win,
            title="Settings",
            style=wx.CAPTION | wx.SYSTEM_MENU | wx.CLOSE_BOX,
        )
        self.InitUi()
        self.Center()

    def InitUi(self):
        panel = wx.Panel(self)
        notebook = wx.Notebook(panel)
        sizer = wx.BoxSizer(wx.VERTICAL)

        tab1_prefs = Prefs(notebook)
        tab2_loc = Location(notebook, self)
        tab3_about = About(notebook)

        notebook.AddPage(tab1_prefs, "Preferences")
        notebook.AddPage(tab2_loc, "Path")
        notebook.AddPage(tab3_about, "About")

        sizer.Add(notebook, flag=wx.EXPAND)
        panel.SetSizer(sizer)
        sizer.Fit(self)


class Prefs(wx.Panel):
    def __init__(self, parent_nb):
        self.parent_nb = parent_nb
        self.pref_dict = dict()
        self.count = 0
        super(Prefs, self).__init__(self.parent_nb)
        self.TabUI()

    def TabUI(self):
        panel = self
        self.sizer = wx.BoxSizer(wx.VERTICAL)
        self.options = [
            "Qualification Information",
            "Tender Details",
            "Bidder Details",
            "Methodology of work",
            "Undertaking(No objection)",
            "Undertaking(On going work)",
            "Undertaking(Bid validity)",
            "Undertaking(Min. cash invest)",
            "Declaration(No near relatives)",
            "Authorization seek reference",
            "Acceptance/Non acceptance of dispute review expert",
            "Information on Litigation History in which the bidder is involved",
            "Proposed sub-contractors and firms involved for construction",
            "Works for which bids are already submitted",
            "Milestone",
        ]
        self.info = wx.StaticText(
            panel,
            label="Here, you can adjust the order in which the documents will be generated.",
        )
        self.option_text = wx.StaticText(panel, label="1")
        self.choice_text = wx.StaticText(panel, label="")
        self.error_text = wx.StaticText(panel, label="")
        self.option = wx.ComboBox(panel, choices=self.options, style=wx.CB_READONLY)
        self.done_icon = wx.StaticBitmap(panel, bitmap=wx.Bitmap("icons/saved.png"))
        self.done_icon.Hide()
        self.save_btn = wx.Button(panel, label="Save preference")
        self.option.Bind(wx.EVT_COMBOBOX, self.onSelect)
        self.save_btn.Bind(wx.EVT_BUTTON, self.onSavePref)

        self.sizer.Add(self.info, flag=wx.TOP | wx.ALIGN_CENTER, border=20)
        self.sizer.Add(self.option_text, flag=wx.ALL | wx.ALIGN_CENTER, border=10)
        self.sizer.Add(self.option, flag=wx.ALL | wx.ALIGN_CENTER, border=10)
        self.sizer.Add(self.choice_text, flag=wx.ALIGN_CENTER | wx.ALL, border=10)
        self.sizer.Add(self.save_btn, flag=wx.ALIGN_CENTER | wx.TOP, border=10)
        self.sizer.Add(self.done_icon, flag=wx.ALIGN_CENTER | wx.ALL, border=20)
        self.sizer.Add(self.error_text, flag=wx.ALL | wx.ALIGN_CENTER, border=20)
        panel.SetSizer(self.sizer)
        self.sizer.Fit(self)

    def onSelect(self, e):
        self.count += 1
        self.error_text.SetLabel("")
        priority = self.option_text.GetLabel()
        if (int(priority)) <= 15:
            choice = e.GetString()
            next_label = str(int(priority) + 1)
            if int(next_label) == 16:
                next_label = "-END-"
            self.option_text.SetLabel(next_label)
            self.options.remove(choice)
            self.option.Clear()
            self.option.AppendItems(self.options)
            self.choice_text.SetLabel(priority + ". " + choice)
            self.sizer.Layout()
            self.pref_dict[priority] = choice
        else:
            self.option_text.SetLabel("-END-")
            self.option.Clear()
            self.sizer.Layout()

    def onSavePref(self, e):
        if self.count < 15:
            self.error_text.SetLabel("Finish selecting all the pages first !")
        else:
            order_pref_db = open("order_pref.pickle", "wb")
            pickle.dump(self.pref_dict, order_pref_db)
            order_pref_db.close()
            self.done_icon.Show()
        self.sizer.Layout()


class Location(wx.Panel):
    def __init__(self, parent_nb, setting_frame):
        self.parent_nb = parent_nb
        self.setting_frame = setting_frame
        super(Location, self).__init__(self.parent_nb, size=(200, 350))
        self.TabUI()

    def TabUI(self):
        panel = self
        self.sizer = wx.BoxSizer(wx.VERTICAL)
        self.browse_btn = wx.Button(panel, label="Browse Path")
        self.info_text = wx.StaticText(
            panel,
            label="Here, you can select the directory where you want your files to be stored once generated.",
        )
        self.path_text = wx.StaticText(panel, label="Selected path appears here. ")
        self.save_btn = wx.Button(panel, label="Save")
        self.saved_icon = wx.StaticBitmap(panel, bitmap=wx.Bitmap("icons/saved.png"))
        self.saved_icon.Hide()
        self.browse_btn.Bind(wx.EVT_BUTTON, self.onClick)
        self.save_btn.Bind(wx.EVT_BUTTON, self.onSave)
        self.sizer.Add(self.info_text, flag=wx.TOP | wx.ALIGN_CENTER, border=30)
        self.sizer.Add(self.browse_btn, flag=wx.ALL | wx.ALIGN_CENTER, border=20)
        self.sizer.Add(self.path_text, flag=wx.ALIGN_CENTER)
        self.sizer.Add(self.save_btn, flag=wx.TOP | wx.ALIGN_CENTER, border=50)
        self.sizer.Add(self.saved_icon, flag=wx.ALL | wx.ALIGN_CENTER, border=10)

        panel.SetSizer(self.sizer)
        self.sizer.Fit(self)

    def onClick(self, e):
        dir_dialog = wx.DirDialog(
            self.setting_frame,
            "Choose Path",
            os.getcwd() + "/",
            wx.DD_DEFAULT_STYLE | wx.DD_DIR_MUST_EXIST,
        )
        dir_dialog.ShowModal()
        self.path = dir_dialog.GetPath() + "/"
        if self.path == "":
            self.path = os.getcwd() + "/"
        self.path_text.SetLabel("Selected path is : ' {} '".format(self.path))
        self.sizer.Layout()

    def onSave(self, e):
        location_db = open("location_pref.pickle", "wb")
        location_dict = dict()
        location_dict["location"] = self.path
        pickle.dump(location_dict, location_db)
        location_db.close()
        self.saved_icon.Show()
        self.sizer.Layout()


class About(wx.Panel):
    def __init__(self, parent_nb):
        self.parent_nb = parent_nb
        super(About, self).__init__(self.parent_nb)
        self.TabUI()

    def TabUI(self):
        panel = self
        sizer = wx.BoxSizer(wx.VERTICAL)
        font = wx.Font(
            13, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD, False
        )
        about_icon = wx.StaticBitmap(panel, bitmap=wx.Bitmap("icons/docImg.png"))
        title_text = wx.StaticText(panel, label="Tender Gen")
        title_text.SetOwnFont(font)
        app_version = wx.StaticText(panel, label="Version 1.0")
        about_text = wx.StaticText(
            panel,
            label="Tender Gen is a software which is used for automating the task of generating Tender related documents.",
        )
        developer_text_1 = wx.StaticText(
            panel, label="Designed and Developed by Surajeet Das."
        )
        developer_text_2 = wx.StaticText(panel, label="CS Engineer.")
        developer_text_3 = wx.StaticText(
            panel, label="Email : das.surajeet97@gmail.com"
        )
        sizer.Add(about_icon, flag=wx.TOP | wx.ALIGN_CENTER, border=10)
        sizer.Add(title_text, flag=wx.TOP | wx.ALIGN_CENTER, border=10)
        sizer.Add(app_version, flag=wx.TOP | wx.ALIGN_CENTER, border=5)
        sizer.Add(about_text, flag=wx.ALL | wx.ALIGN_CENTER, border=40)
        sizer.Add(developer_text_1, flag=wx.ALIGN_CENTER)
        sizer.Add(developer_text_2, flag=wx.TOP | wx.ALIGN_CENTER, border=5)
        sizer.Add(developer_text_3, flag=wx.ALL | wx.ALIGN_CENTER, border=10)

        panel.SetSizer(sizer)
        sizer.Fit(self)


def main():
    app_instance = wx.App()
    app_frame = TenderGenApp(None, "Tender Gen Software")
    app_frame.Show()
    app_instance.MainLoop()


if __name__ == "__main__":
    main()
