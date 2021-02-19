# -*- coding: utf-8 -*-
##############################################################################
#                                                                            #
# Part of Caret IT Solutions Pvt. Ltd. (Website: www.caretit.com).           #
# See LICENSE file for full copyright and licensing details.                 #
#                                                                            #
##############################################################################

from odoo import api, fields, models, _
from odoo.exceptions import UserError
from collections import defaultdict
from datetime import datetime
import xlwt
import io
from odoo.tools.float_utils import float_round
from datetime import datetime
from dateutil.relativedelta import relativedelta
from odoo.tools.misc import formatLang

class ProfitLossCustomReport(models.TransientModel):
    _name = 'profit.loss.custom.report'
    _description = "Profit And Loss Cutom Report"


    date_from = fields.Date(string="From Date")
    date_to = fields.Date(string="To Date")
    account_ids = fields.Many2many('account.account', string='Accounts')
    account_income_percentage = fields.Boolean(string= 'Show Income Percentage', default=False)
    income_percentage = fields.Selection([('repeatnone','Repeat None'),('repeattag','Repeat Tag Wise')], string='income percentage', default='repeatnone')
    dimension_wise_project = fields.Selection([('none','None'),('month','Month Wise'),('dimension', 'Dimension'),('year','Year Wise')],
                                              string='Dimension',
                                              default='none')
    projectwise = fields.Selection([('project', 'Project')],string='Project',default='project')
    analytic_account_ids = fields.Many2many('account.analytic.account', string='Analytic Accounts')


    @api.model
    def format_value(self, amount, currency=False, blank_if_zero=False):
        ''' Format amount to have a monetary display (with a currency symbol).
        E.g: 1000 => 1000.0 $

        :param amount:          A number.
        :param currency:        An optional res.currency record.
        :param blank_if_zero:   An optional flag forcing the string to be empty if amount is zero.
        :return:                The formatted amount as a string.
        '''
        currency_id = currency or self.env.company.currency_id
        if currency_id.is_zero(amount):
            if blank_if_zero:
                return ''
            # don't print -0.0 in reports
            amount = abs(amount)

        if self.env.context.get('no_format'):
            return amount

        return formatLang(self.env, amount, currency_obj=currency_id)

    def print_report_trial_balance(self):
        if self.date_from >= self.date_to:
            raise UserError(_("Start Date is greater than or equal to End Date."))
        datas = {'form': self.read()[0],
                 'get_trial_balance': self.get_trial_balance_detail()
            }
        return self.env.ref('account_trial_balance.action_report_trial_balance').report_action([], data=datas)

    def print_report(self):
      if self.date_from >= self.date_to:
          raise UserError(_("Start Date is greater than or equal to End Date."))
      datas = {'form': self.read()[0],
               'get_profit_loss': self.get_profit_and_loss_detail()
          }
      return self.env.ref('profit_and_loss_custom_report.action_report_profit_loss').report_action([], data=datas)

    def get_profit_and_loss_detail(self):
        CompanyImage = self.env.company.logo
        dateFrom = self.date_from
        dateTo = self.date_to
        MoveLineIds = []
        Vals = {}
        mainDict = []
        new_list = []
        account_names = []
        AllAccounts = self.account_ids
        FilteredAccountIds = AllAccounts.filtered(lambda a: a.temp_account_report)
        AccountIds = FilteredAccountIds.ids
        if not AccountIds:
            AccountIds = AllAccounts.ids
        AllAnalyticAccounts = self.analytic_account_ids
        FilteredAnalyticAccountIds = AllAnalyticAccounts.filtered(lambda a: a.temp_analytics_report)
        AnalyticAccountIds = FilteredAnalyticAccountIds
        if not AnalyticAccountIds:
            AnalyticAccountIds = AllAnalyticAccounts
        Status = ['posted']
        Projectwise = self.dimension_wise_project
        for Account in self.env['account.account'].browse(AccountIds):
            Balance = 0.0
            self.env.cr.execute("""
                SELECT aml.date as date,
                       aml.debit as debit,
                       aml.credit as credit,
                       a.code as code,
                       a.name as acc_name,
                       at.name as acc_type,
                       aa.name as analytic,
                       aml.id as movelineid
                FROM account_move_line aml
                LEFT JOIN account_move am ON (am.id=aml.move_id)
                LEFT JOIN account_analytic_account aa ON (aa.id=aml.analytic_account_id)
                LEFT JOIN account_account a ON (a.id=aml.account_id)
                LEFT JOIN account_account_type at ON (at.id=a.user_type_id)
                WHERE (aml.date >= %s) AND
                    (aml.date <= %s) AND
                    (aml.account_id in %s) AND
                    (am.state in %s) ORDER BY aml.date""",
                (str(dateFrom) + ' 00:00:00', str(dateTo) + ' 23:59:59', tuple([Account.id]), tuple(Status),))
            MoveLineIds = self.env.cr.fetchall()
            if MoveLineIds:
                for ml in MoveLineIds:
                    date = ml[0]
                    acount_debit = ml[1]
                    account_credit = ml[2]
                    account_code = ml[3]
                    account_name = ml[4]
                    account_type = ml[5]
                    analytic_account_id = ml[6]
                    Balance = 0.0
                    Balance = Balance + (acount_debit - account_credit)
                    Vals = {'account_code':account_code,
                            'account_name':account_name,
                            'balance': Balance or 0.0,
                            'percentage': 0.0,
                            'account_type':account_type,
                            'account_debit':acount_debit,
                            'account_credit':account_credit,
                            'analytic_account_id':analytic_account_id,
                            'date':date,
                            }
                    mainDict.append(Vals)

        for i in range(0,len(mainDict)):
            if mainDict[i]['account_name'] not in account_names:
                new_list.append(mainDict[i])
                account_names.append(mainDict[i]['account_name'])
            else:
                for j in range(0,len(new_list)):
                    if mainDict[i]['account_name'] == new_list[j]['account_name']:
                        new_list[j]['balance'] = new_list[j]['balance'] + mainDict[i]['balance']
        return new_list


    @api.model
    def default_get(self, fields):
        vals = super(ProfitLossCustomReport, self).default_get(fields)
        ac_ids = self.env['account.account'].search([])
        analytic_ids = self.env['account.analytic.account'].search([])
        self.env.cr.execute('update account_account set temp_account_report=False')
        self.env.cr.execute('update account_analytic_account set temp_analytics_report=False')
        if 'account_ids' in fields and not vals.get('account_ids') and ac_ids:
            iids = []
            for ac_id in ac_ids:
                iids.append(ac_id.id)
            vals['account_ids'] = [(6, 0, iids)]
        if 'analytic_account_ids' in fields and not vals.get('analytic_account_ids') and analytic_ids:
            aniids = []
            for ana_ac in analytic_ids:
                aniids.append(ana_ac.id)
            vals['analytic_account_ids'] = [(6, 0, aniids)]
        return vals
       

    def profit_and_loss_export_excel(self):
        CompanyImage = self.env.company.logo
        dateFrom = self.date_from
        dateTo = self.date_to
        MoveLineIds = []
        Vals = {}
        mainDict = []
        account_list = []
        main_list = []
        first_list = []
        AnalyticVals = []
        AllAccounts = self.account_ids
        FilteredAccountIds = AllAccounts.filtered(lambda a: a.temp_account_report)
        AccountIds = FilteredAccountIds.ids
        if not AccountIds:
            AccountIds = AllAccounts.ids
        AllAnalyticAccounts = self.analytic_account_ids
        FilteredAnalyticAccountIds = AllAnalyticAccounts.filtered(lambda a: a.temp_analytics_report)
        AnalyticAccountIds = FilteredAnalyticAccountIds
        if not AnalyticAccountIds:
            AnalyticAccountIds = AllAnalyticAccounts
        Status = ['posted']
        Projectwise = self.dimension_wise_project
        for Account in self.env['account.account'].browse(AccountIds):
            Balance = 0.0
            self.env.cr.execute("""
                SELECT aml.date as date,
                       aml.debit as debit,
                       aml.credit as credit,
                       a.code as code,
                       a.name as acc_name,
                       at.name as acc_type,
                       aa.name as analytic,
                       aml.id as movelineid
                FROM account_move_line aml
                LEFT JOIN account_move am ON (am.id=aml.move_id)
                LEFT JOIN account_analytic_account aa ON (aa.id=aml.analytic_account_id)
                LEFT JOIN account_account a ON (a.id=aml.account_id)
                LEFT JOIN account_account_type at ON (at.id=a.user_type_id)
                WHERE (aml.date >= %s) AND
                    (aml.date <= %s) AND
                    (aml.account_id in %s) AND
                    (am.state in %s) ORDER BY aml.date""",
                (str(dateFrom) + ' 00:00:00', str(dateTo) + ' 23:59:59', tuple([Account.id]), tuple(Status),))
            MoveLineIds = self.env.cr.fetchall()
            if MoveLineIds:
                for ml in MoveLineIds:
                    date = ml[0]
                    acount_debit = ml[1]
                    account_credit = ml[2]
                    account_code = ml[3]
                    account_name = ml[4]
                    account_type = ml[5]
                    analytic_account_id = ml[6]
                    Balance = 0.0
                    Balance = Balance + (acount_debit - account_credit)
                    Vals = {'account_code':account_code,
                            'account_name':account_name,
                            'balance': Balance or 0.0,
                            'percentage': 0.0,
                            'account_type':account_type,
                            'account_debit':acount_debit,
                            'account_credit':account_credit,
                            'analytic_account_id':analytic_account_id,
                            'date':date,
                            }
                    mainDict.append(Vals)
        if Projectwise == 'dimension':
            # second_list = []
            # news_list = []
            for i in range(0,len(mainDict)):
                if (mainDict[i]['account_name'],mainDict[i]['analytic_account_id']) not in account_list:
                    main_list.append({
                                      'account_name':mainDict[i]['account_name'],
                                      'analytic_account_id':mainDict[i]['analytic_account_id'],
                                      'account_type':mainDict[i]['account_type'],
                                      'debit': mainDict[i]['account_debit'],
                                      'credit': mainDict[i]['account_credit'],
                                      'balance': mainDict[i]['account_debit'] - mainDict[i]['account_credit'] or 00.00,
                                      })
                    account_list.append((mainDict[i]['account_name'],mainDict[i]['analytic_account_id']))
                else:
                    first_list.append({
                                      'account_name':mainDict[i]['account_name'],
                                      'analytic_account_id':mainDict[i]['analytic_account_id'],
                                      'account_type':mainDict[i]['account_type'],
                                      'debit': mainDict[i]['account_debit'],
                                      'credit': mainDict[i]['account_credit'],
                                      'balance': mainDict[i]['account_debit'] - mainDict[i]['account_credit'] or 00.00,
                                      })
            for j in range(0,len(main_list)):
                for k in range(0,len(first_list)):
                    for ana in AnalyticAccountIds:
                        if first_list[k]['account_name'] == main_list[j]['account_name'] and first_list[k]['analytic_account_id'] == main_list[j]['analytic_account_id'] and ana.name == main_list[j]['analytic_account_id']:
                            main_list[j]['debit'] =  main_list[j]['debit'] + first_list[k]['debit']
                            main_list[j]['credit'] = main_list[j]['credit'] + first_list[k]['credit']
                            main_list[j]['balance'] = main_list[j]['debit'] - main_list[j]['credit']

        if Projectwise == 'month':
            for i in range(0,len(mainDict)):
                if (mainDict[i]['account_name'],mainDict[i]['date'].strftime("%b %y")) not in account_list:
                    main_list.append({
                                      'account_name':mainDict[i]['account_name'],
                                      'debit': mainDict[i]['account_debit'],
                                      'credit': mainDict[i]['account_credit'],
                                      'balance': mainDict[i]['account_debit'] - mainDict[i]['account_credit'] or 00.00,
                                      'account_type':mainDict[i]['account_type'],
                                      'month': mainDict[i]['date'].strftime("%b %y")
                                      })
                    account_list.append((mainDict[i]['account_name'],mainDict[i]['date'].strftime("%b %y")))        
                else:
                    first_list.append({
                                      'account_name':mainDict[i]['account_name'],
                                      'debit': mainDict[i]['account_debit'],
                                      'credit': mainDict[i]['account_credit'],
                                      'balance': mainDict[i]['account_debit'] - mainDict[i]['account_credit'] or 00.00,
                                      'account_type':mainDict[i]['account_type'],
                                      'month': mainDict[i]['date'].strftime("%b %y")
                                      })
            if mainDict:
                for j in range(0,len(main_list)):
                    for k in range(0,len(first_list)):
                        if first_list[k]['account_name'] == main_list[j]['account_name'] and first_list[k]['month'] == main_list[j]['month']:
                            main_list[j]['debit'] =  main_list[j]['debit'] + first_list[k]['debit']
                            main_list[j]['credit'] = main_list[j]['credit'] + first_list[k]['credit']
                            main_list[j]['balance'] = main_list[j]['debit'] - main_list[j]['credit']
        if Projectwise == 'year':
            for i in range(len(mainDict)):
                if (mainDict[i]['account_name'],mainDict[i]['date'].strftime("%Y")) not in account_list:
                        main_list.append({
                                          'account_name':mainDict[i]['account_name'],
                                          'debit': mainDict[i]['account_debit'],
                                          'credit': mainDict[i]['account_credit'],
                                          'balance': mainDict[i]['account_debit'] - mainDict[i]['account_credit'] or 00.00,
                                          'account_type':mainDict[i]['account_type'],
                                          'year': mainDict[i]['date'].strftime("%Y")
                                          })
                        account_list.append((mainDict[i]['account_name'],mainDict[i]['date'].strftime("%Y")))        
                else:
                    first_list.append({
                                      'account_name':mainDict[i]['account_name'],
                                      'debit': mainDict[i]['account_debit'],
                                      'credit': mainDict[i]['account_credit'],
                                      'balance': mainDict[i]['account_debit'] - mainDict[i]['account_credit'] or 00.00,
                                      'account_type':mainDict[i]['account_type'],
                                      'year': mainDict[i]['date'].strftime("%Y")
                                      })
            if mainDict:
                for j in range(0,len(main_list)):
                    for k in range(0,len(first_list)):
                        if first_list[k]['account_name'] == main_list[j]['account_name'] and first_list[k]['year'] == main_list[j]['year']:
                            main_list[j]['debit'] =  main_list[j]['debit'] + first_list[k]['debit']
                            main_list[j]['credit'] = main_list[j]['credit'] + first_list[k]['credit']
                            main_list[j]['balance'] = main_list[j]['debit'] - main_list[j]['credit']

        import base64
        dateFrom = self.date_from
        dateTo = self.date_to
        filename = 'Profit And Loss.xls'
        form_name = 'Profit Loss Between ' + str(dateFrom) + ' to ' + str(dateTo)
        workbook = xlwt.Workbook()
        style = xlwt.XFStyle()
        tall_style = xlwt.easyxf('font:height 720;') # 36pt
        # Create a font to use with the style
        font = xlwt.Font()
        font.name = 'Times New Roman'
        font.bold = True
        font.height = 250
        style.font = font
        xlwt.add_palette_colour("custom_colour", 0x21)
        workbook.set_colour_RGB(0x21, 105, 105, 105)

        xlwt.add_palette_colour("dark_blue", 0x3A)
        workbook.set_colour_RGB(0x3A, 0,0,139)  

        xlwt.add_palette_colour("gainsboro", 0x15)
        workbook.set_colour_RGB(0x15,205,205,205)  

        worksheet = workbook.add_sheet("Profit And Loss", cell_overwrite_ok=True)
        worksheet.show_grid = False

        styleheader = xlwt.easyxf('font: bold 1, colour black, height 300;')
        stylecolumnheader = xlwt.easyxf('font: bold 1, colour black, height 200;pattern: pattern solid, fore_colour gainsboro')
        linedata = xlwt.easyxf('align: horiz right;')
        alinedata = xlwt.easyxf('align: horiz left;','#,##.00')
        stylecolaccount = xlwt.easyxf('font: bold 1, colour white, height 200; \
                                      pattern: pattern solid, fore_colour dark_blue; \
                                      align: vert centre, horiz centre;')
        analytic_st_col = xlwt.easyxf('font: bold 1, colour black, height 200; \
                                    pattern: pattern solid, fore_colour gray_ega; \
                                    align: vert centre, horiz centre;')
        general = xlwt.easyxf('font: bold 1, colour black, height 210;')
        dateheader = xlwt.easyxf('font: bold 1, colour black, height 200;')
        maintotal = xlwt.easyxf('font: bold 1, colour black, height 200;')
        finaltotalheader = xlwt.easyxf('pattern: fore_color white; font: bold 1, colour black; align: horiz right;')
        rightfont = xlwt.easyxf('pattern: fore_color white; align: horiz right;')
        floatstyle = xlwt.easyxf("align: horiz right;","#,##0.00")
        finaltotalheaderbold = xlwt.easyxf("pattern: fore_color white; font: bold 1, colour black;")
        accountnamestyle = xlwt.easyxf('font: bold 1, colour green, height 200;')
        mainheaders = xlwt.easyxf('pattern: fore_color white; font: bold 1, colour dark_blue; align: horiz left;')
        mainheader = xlwt.easyxf('pattern: pattern solid, fore_colour gainsboro; \
                                 font: bold 1, colour dark_blue; align: horiz left;')
        mainheaderexpense = xlwt.easyxf('pattern: pattern solid, fore_colour gainsboro; \
                                 font: bold 1, colour dark_blue; align: horiz left;borders: top_color black,\
                              top double;')
        netmainheader = xlwt.easyxf('pattern: pattern solid, fore_colour gainsboro; \
                                 font: bold 1, colour dark_blue; align: horiz right;borders: bottom_color black,\
                              bottom double;')
        netmainheaders = xlwt.easyxf('pattern: pattern solid, fore_colour gainsboro; \
                                 font: bold 1, colour dark_blue; align: horiz left;borders: bottom_color black,\
                              bottom double;')
        mainheaderline = xlwt.easyxf("pattern: pattern solid, fore_colour gainsboro; \
                                 font: bold 1, colour dark_blue; align: horiz right;")
        mainheaderdata = xlwt.easyxf("pattern: fore_color white; font: bold 1, colour dark_blue; align: horiz right;")
        mainheaderdatas = xlwt.easyxf("pattern: fore_color white; font: bold 1, colour dark_blue; align: horiz right;")
        zero_col = worksheet.col(0)
        zero_col.width = 236 * 22
        first_col = worksheet.col(1)
        first_col.width = 236 * 40
        second_col = worksheet.col(2)
        second_col.width = 236 * 40
        third_col = worksheet.col(3)
        third_col.width = 236 * 25
        fourth_col = worksheet.col(4)
        fourth_col.width = 236 * 20
        fifth_col = worksheet.col(5)
        fifth_col.width = 236 * 20
        sixth_col = worksheet.col(6)
        sixth_col.width = 236 * 20
        seventh_col = worksheet.col(7)
        seventh_col.width = 236 * 20
        #HEADER
        worksheet.row(4).height_mismatch = True
        worksheet.row(4).height = 360
        
        worksheet.write_merge(0, 1, 2, 5, self.env.company.name,styleheader)
        worksheet.write_merge(2, 2, 2, 5, 'Profit & Loss',general)
        headerstring = 'From :' + str(self.date_from.strftime('%d %b %Y') or '') + ' To :' + str(self.date_to.strftime('%d %b %Y') or '')
        worksheet.write_merge(3, 3, 2, 5, headerstring,dateheader)
        
        #SUB-HEADER
        row = 4
        ColIndexes = {}
        worksheet.write(row, 0, 'Account  Code', stylecolaccount)
        worksheet.write(row, 1, 'Account Name', stylecolaccount)
        worksheet.write(row, 2, 'Balance', stylecolaccount)
        worksheet.write(row, 3, 'Percentage', stylecolaccount)
        calc = 5
        col = 4
        colc = 4

        if self.dimension_wise_project == 'dimension':
            for analytic in AnalyticAccountIds:
                dictval = {analytic.name:col}
                ColIndexes.update(dictval)
                dyna_col = worksheet.col(col)
                dyna_col.width = 236 * 20
                worksheet.write(row, col, analytic.name, analytic_st_col)
                colc = col
                col+=1
                calc+=1
        elif self.dimension_wise_project == 'month':
            cur_date = datetime.strptime(str(self.date_from), '%Y-%m-%d').date()
            end = datetime.strptime(str(self.date_to), '%Y-%m-%d').date()
            while cur_date < end:
                cur_date_strf = str(cur_date.strftime('%b %y') or '')
                cur_date += relativedelta(months=1)
                dictval = {cur_date_strf : col }
                ColIndexes.update(dictval)
                dyna_col = worksheet.col(col)
                dyna_col.width = 236 * 20
                worksheet.write(row, col, cur_date_strf, analytic_st_col)
                colc = col
                col+=1
                calc+=1
        elif self.dimension_wise_project == 'year':
            cur_date = datetime.strptime(str(self.date_from), '%Y-%m-%d').date()
            end = datetime.strptime(str(self.date_to), '%Y-%m-%d').date()
            while cur_date <= end:
                cur_date_strf = str(cur_date.strftime('%Y') or '')
                dictval = {cur_date_strf : col }
                cur_date += relativedelta(years=1)
                ColIndexes.update(dictval)
                dyna_col = worksheet.col(col)
                dyna_col.width = 236 * 20
                worksheet.write(row, col, cur_date_strf, analytic_st_col)
                colc = col
                col+=1
                calc+=1
        cols = col

        row = 5
        FinalDict = {}
        new_list = []
        account_name = []
        defaultpercentage = 100.00
        OpPercentage = 00
        OinPercentage = 00
        CorPercentage = 00
        ExPercentage = 00
        DepPercentage = 00
        OperatingIncome = 00
        OtherIncome = 00
        CostOfRevenue = 00
        TotalExpenses = 00
        TotalIncomeHeading = 00
        Depreciation = 00
        Expenses = 00
        TotalIncome = 00
        GrossProfit = 00
        NetProfit = 00
        Totalnetprofitloss = 00

        for i in range(0,len(mainDict)):
            worksheet.row(row).height_mismatch = True
            worksheet.row(row).height = 310
            if mainDict[i]['account_name'] not in account_name:
                new_list.append(mainDict[i])
                account_name.append(mainDict[i]['account_name'])
            else:
                for j in range(0,len(new_list)):
                    if mainDict[i]['account_name'] == new_list[j]['account_name']:
                        new_list[j]['balance'] = new_list[j]['balance'] + mainDict[i]['balance']

    
        for k in range(0,len(new_list)):
            try :
                if  new_list[k]['account_type'] == "Income" :
                    OperatingIncome = OperatingIncome + abs(new_list[k]['balance'])
                elif new_list[k]['account_type'] == "Other Income":
                    OtherIncome = OtherIncome + new_list[k]['balance']
                elif new_list[k]['account_type'] == "Cost of Revenue":
                    CostOfRevenue += new_list[k]['balance']
                elif new_list[k]['account_type'] == "Depreciation":
                    Depreciation += new_list[k]['balance']
                elif new_list[k]['account_type'] == "Expenses":
                    Expenses += new_list[k]['balance']
                # ExpensesHeading = Expenses + Depreciation
                GrossProfit = abs(OperatingIncome) - CostOfRevenue
                OpPercentage = round((OperatingIncome * 100 / GrossProfit),1)
                GrossProfitPercentage = GrossProfit *100 / OperatingIncome
                ExpensePercentage = round((TotalExpenses * 100 /OperatingIncome ),1)
                CorPercentage = round((CostOfRevenue * 100 / OperatingIncome),1)
                # ExPercentage = round((Expenses * 100 / GrossProfit),1)
                DepPercentage = round((Depreciation * 100 / OperatingIncome),1)
                TotalIncome = GrossProfit + abs(OtherIncome)

                TotalExpenses = Expenses + Depreciation
                NetProfit = TotalIncome - TotalExpenses
               
            except :
                pass

        worksheet.write(row, 0,'Income', style = mainheader)
        worksheet.write(row, 1,'', style = mainheader)
        worksheet.write(row, 2,'', style = mainheader)
        worksheet.write(row, 3,'', style = mainheader)
        for i in range(4,100):
          if i != col:
              worksheet.write(row, i,'',style = mainheader)
          elif i == col:
              break
        row +=1
        worksheet.write(row, 0,'', style = mainheaders)
        worksheet.write(row, 1,'Gross Profit', style = mainheaders)
        worksheet.write(row, 2,'', style = mainheaderdata)
        worksheet.write(row, 3,'', style = mainheaderdata)
        for i in range(4,100):
          if i != col:
              worksheet.write(row, i,'',style = mainheaderdatas)
          elif i == col:
              break
        row +=1
        worksheet.write(row, 0,'', style = mainheaders)
        worksheet.write(row,1, 'Operating Income', style = mainheaders)
        worksheet.write(row,2, '', style = mainheaderdata )
        worksheet.write(row,3, '', style = mainheaderdatas)
        for i in range(4,100):
          if i != col:
              worksheet.write(row, i,'',style = mainheaderdatas)
          elif i == col:
              break

        finalres_list = []
        netbalance_list = []
        third_income_lists = []
        third_expense_lists = []
        totalincomecolumn = [] 
        thirdincomelists = []
        totalcostcolumn = [] 
        thirdcostlists = [] 
        totalothercolumn = [] 
        thirdotherlists = [] 
        totalexpensecolumn = [] 
        thirdexpenselists = [] 
        totaldepriciationcolumn = [] 
        thirddepriciationlists = [] 
        finalgross_list =[] 
        finalincome_list = [] 
        finalexpense_list =[]

        if Projectwise == 'dimension':
            a3 = ''
            res2 =''
            column1 = []
            news_list = []
            second_list = []
            check_list = []
            AnalyticAccountsId = [i.id for i in AnalyticAccountIds]
            ana_id = self.env['account.analytic.account'].browse(AnalyticAccountsId)
            ac_names = [i.name for i in ana_id]
            for val in main_list:
                if list(val.values())[1] is None:
                    continue                                                                                          
                else:
                    check_list.append(val)

            for i in range(0,len(check_list)):       
                if check_list[i]['account_name'] not in second_list:
                    news_list.append(check_list[i])
                    second_list.append(check_list[i]['account_name'])

            for k in range(0,len(news_list)):
                for data in range(0,len(check_list)):
                    if news_list[k]['account_name'] == check_list[data]['account_name']:
                        column1.append({check_list[data]['analytic_account_id']:check_list[data]['balance']})
                        a1 = [(list(c.keys())[0]) for c in column1]
                        res = column1 + [{i:000.0} for i in ac_names if i not in a1]
                        res2 = sorted(res, key = lambda ele: ac_names.index(list(ele.keys())[0]))
                        news_list[k]['columns'] = res2
                        news_list[k]['caret_options'] = 'account.account'
                    else:
                        column1.clear()

            for s in range(0,len(news_list)):

                if news_list[s]['account_type'] == "Income" :
                    totalincomecolumn = news_list[s]['columns']
                    listd = [list(c.values())[0] for c in totalincomecolumn]
                    thirdincomelists.append(listd)

                if news_list[s]['account_type'] == "Cost of Revenue":
                    totalcostcolumn = news_list[s]['columns']
                    listd = [list(c.values())[0] for c in totalcostcolumn]
                    thirdcostlists.append(listd)

                if news_list[s]['account_type'] == "Other Income" :
                    totalothercolumn = news_list[s]['columns']
                    listd = [list(c.values())[0] for c in totalothercolumn]
                    thirdotherlists.append(listd)

                if news_list[s]['account_type'] == "Expenses" :
                    totalexpensecolumn = news_list[s]['columns']
                    listd = [list(c.values())[0] for c in totalexpensecolumn]
                    thirdexpenselists.append(listd)

                if news_list[s]['account_type'] == "Depreciation" :
                    totaldepriciationcolumn = news_list[s]['columns']
                    listd = [list(c.values())[0] for c in totaldepriciationcolumn]
                    thirddepriciationlists.append(listd)

            thirdincomelist = [sum(i) for i in zip(*thirdincomelists)]
            thirdcostlist = [sum(i) for i in zip(*thirdcostlists)]
            thirdotherlist = [sum(i) for i in zip(*thirdotherlists)]
            thirdexpenselist = [sum(i) for i in zip(*thirdexpenselists)]
            thirddepriciationlist = [sum(i) for i in zip(*thirddepriciationlists)]

            for i in range(0, len(thirdincomelist)):
                if thirdincomelist and thirdcostlist:
                    finalgross_list.append(thirdincomelist[i] - thirdcostlist[i])
                elif thirdincomelist:
                    finalgross_list.append(thirdincomelist[i])
                elif thirdcostlist:
                    finalgross_list.append(thirdincomelist[i])

            for i in range(0,len(finalgross_list)):
                if finalgross_list and thirdotherlist[i]:
                    finalincome_list.append(finalgross_list[i] + thirdotherlist[i])
                elif finalgross_list:
                    finalincome_list.append(finalgross_list[i])
                elif thirdotherlist:
                    finalincome_list.append(thirdotherlist[i])

            for i in range(0, len(thirdexpenselist)):
                if thirdexpenselist and thirddepriciationlist:
                    finalexpense_list.append(thirdexpenselist[i] - thirddepriciationlist[i])
                elif thirdexpenselist:
                    finalexpense_list.append(thirdexpenselist[i])
                elif thirddepriciationlist:
                    finalexpense_list.append(thirddepriciationlist[i])

            for i in range(0, len(finalincome_list)):
                netbalance_list.append(finalincome_list[i] - finalexpense_list[i])

        if Projectwise == 'month':
            a1 = ''
            res2 =''
            fetch_monthwise_data = []
            news_list = []
            second_list = []
            column1 = []

            cur_date = dateFrom
            end = dateTo
            while cur_date < end:
                cur_date_strf = str(cur_date.strftime('%b %y') or '')
                cur_date += relativedelta(months=1)
                fetch_monthwise_data.append(cur_date_strf)

            for i in range(0,len(main_list)):
                if main_list[i]['account_name'] not in second_list:
                    news_list.append(main_list[i])
                    second_list.append(main_list[i]['account_name'])

            for j in range(0,len(news_list)):
                for k in range(0,len(main_list)):
                    if news_list[j]['account_name'] == main_list[k]['account_name']:
                        column1.append({main_list[k]['month']:main_list[k]['balance']})
                        a1 = [(list(c.keys())[0]) for c in column1]
                        res = column1 + [{i:000.0} for i in fetch_monthwise_data if i not in a1]
                        res2 = sorted(res, key = lambda ele: fetch_monthwise_data.index(list(ele.keys())[0]))
                        news_list[j]['columns'] = res2
                        news_list[j]['caret_options'] = 'account.account'
                        
                    else:
                       column1.clear()

            for s in range(0,len(news_list)):

                if news_list[s]['account_type'] == "Income" :
                    totalincomecolumn = news_list[s]['columns']
                    listd = [list(c.values())[0] for c in totalincomecolumn]
                    thirdincomelists.append(listd)

                if news_list[s]['account_type'] == "Cost of Revenue":
                    totalcostcolumn = news_list[s]['columns']
                    listd = [list(c.values())[0] for c in totalcostcolumn]
                    thirdcostlists.append(listd)

                if news_list[s]['account_type'] == "Other Income" :
                    totalothercolumn = news_list[s]['columns']
                    listd = [list(c.values())[0] for c in totalothercolumn]
                    thirdotherlists.append(listd)

                if news_list[s]['account_type'] == "Expenses" :
                    totalexpensecolumn = news_list[s]['columns']
                    listd = [list(c.values())[0] for c in totalexpensecolumn]
                    thirdexpenselists.append(listd)

                if news_list[s]['account_type'] == "Depreciation":
                    totaldepriciationcolumn = news_list[s]['columns']
                    listd = [list(c.values())[0] for c in totaldepriciationcolumn]
                    thirddepriciationlists.append(listd)

            thirdincomelist = [sum(i) for i in zip(*thirdincomelists)]
            thirdcostlist = [sum(i) for i in zip(*thirdcostlists)]
            thirdotherlist = [sum(i) for i in zip(*thirdotherlists)]
            thirdexpenselist = [sum(i) for i in zip(*thirdexpenselists)]
            thirddepriciationlist = [sum(i) for i in zip(*thirddepriciationlists)]

            for i in range(0, len(thirdincomelist)):
                if thirdincomelist and thirdcostlist:
                    finalgross_list.append(thirdincomelist[i] - thirdcostlist[i])
                elif thirdincomelist:
                    finalgross_list.append(thirdincomelist[i])
                elif thirdcostlist:
                    finalgross_list.append(thirdincomelist[i])

            for i in range(0,len(finalgross_list)):
                if finalgross_list and thirdotherlist[i]:
                    finalincome_list.append(finalgross_list[i] + thirdotherlist[i])
                elif finalgross_list:
                    finalincome_list.append(finalgross_list[i])
                elif thirdotherlist:
                    finalincome_list.append(thirdotherlist[i])

            for i in range(0, len(thirdexpenselist)):
                if thirdexpenselist and thirddepriciationlist:
                    finalexpense_list.append(thirdexpenselist[i] - thirddepriciationlist[i])
                elif thirdexpenselist:
                    finalexpense_list.append(thirdexpenselist[i])
                elif thirddepriciationlist:
                    finalexpense_list.append(thirddepriciationlist[i])
                    
            for i in range(0, len(finalincome_list)): 
                netbalance_list.append(finalincome_list[i] + finalexpense_list[i])
          
        if Projectwise == 'year':
            a1 = ''
            res2 =''
            fetch_yearwise_data = []
            news_list = []
            second_list = []
            column1 = []

            for k,v in ColIndexes.items():
                fetch_yearwise_data.append(k)

            for i in range(0,len(main_list)):
                if main_list[i]['account_name'] not in second_list:
                    news_list.append(main_list[i])
                    second_list.append(main_list[i]['account_name'])

            for j in range(0,len(news_list)):
                for k in range(0,len(main_list)):
                    if news_list[j]['account_name'] == main_list[k]['account_name']:
                        column1.append({main_list[k]['year']:main_list[k]['balance']})
                        a1 = [(list(c.keys())[0]) for c in column1]
                        res = column1 + [{i:000.0} for i in fetch_yearwise_data if i not in a1]
                        res2 = sorted(res, key = lambda ele: fetch_yearwise_data.index(list(ele.keys())[0]))
                        news_list[j]['columns'] = res2
                        news_list[j]['caret_options'] = 'account.account'
                        
                    else:
                       column1.clear()

            for s in range(0,len(news_list)):

                if news_list[s]['account_type'] == "Income" :
                    totalincomecolumn = news_list[s]['columns']
                    listd = [list(c.values())[0] for c in totalincomecolumn]
                    thirdincomelists.append(listd)

                if news_list[s]['account_type'] == "Cost of Revenue":
                    totalcostcolumn = news_list[s]['columns']
                    listd = [list(c.values())[0] for c in totalcostcolumn]
                    thirdcostlists.append(listd)

                if news_list[s]['account_type'] == "Other Income" :
                    totalothercolumn = news_list[s]['columns']
                    listd = [list(c.values())[0] for c in totalothercolumn]
                    thirdotherlists.append(listd)

                if news_list[s]['account_type'] == "Expenses" :
                    totalexpensecolumn = news_list[s]['columns']
                    listd = [list(c.values())[0] for c in totalexpensecolumn]
                    thirdexpenselists.append(listd)

                if news_list[s]['account_type'] == "Depreciation":
                    totaldepriciationcolumn = news_list[s]['columns']
                    listd = [list(c.values())[0] for c in totaldepriciationcolumn]
                    thirddepriciationlists.append(listd)

            thirdincomelist = [sum(i) for i in zip(*thirdincomelists)]
            thirdcostlist = [sum(i) for i in zip(*thirdcostlists)]
            thirdotherlist = [sum(i) for i in zip(*thirdotherlists)]
            thirdexpenselist = [sum(i) for i in zip(*thirdexpenselists)]
            thirddepriciationlist = [sum(i) for i in zip(*thirddepriciationlists)]

            for i in range(0, len(thirdincomelist)):
                if thirdincomelist and thirdcostlist:
                    finalgross_list.append(thirdincomelist[i] - thirdcostlist[i])
                elif thirdincomelist:
                    finalgross_list.append(thirdincomelist[i])
                elif thirdcostlist:
                    finalgross_list.append(thirdincomelist[i])

            for i in range(0,len(finalgross_list)):
                if finalgross_list and thirdotherlist[i]:
                    finalincome_list.append(finalgross_list[i] + thirdotherlist[i])
                elif finalgross_list:
                    finalincome_list.append(finalgross_list[i])
                elif thirdotherlist:
                    finalincome_list.append(thirdotherlist[i])

            for i in range(0, len(thirdexpenselist)):
                if thirdexpenselist and thirddepriciationlist:
                    finalexpense_list.append(thirdexpenselist[i] - thirddepriciationlist[i])
                elif thirdexpenselist:
                    finalexpense_list.append(thirdexpenselist[i])
                elif thirddepriciationlist:
                    finalexpense_list.append(thirddepriciationlist[i])

            for i in range(0, len(finalincome_list)): 
                netbalance_list.append(finalincome_list[i] + finalexpense_list[i])

        third_income_list = []

        for k in range(0,len(new_list)):
            if new_list[k]['account_type'] == "Income":
                row+=1
                try:
                    percentage = round(((new_list[k]['balance'] * 100) / OperatingIncome),1)
                    worksheet.write(row, 0, new_list[k]['account_code'],alinedata)
                    worksheet.write(row, 1, new_list[k]['account_name'],alinedata)
                    worksheet.write(row, 2, abs(new_list[k]['balance']),floatstyle)
                    worksheet.write(row, 3, abs(percentage),floatstyle)
                    if Projectwise == 'dimension':
                        mainlist_position={}
                        for main_data in range(len(main_list)):
                            if main_list[main_data]['account_name'] == new_list[k]['account_name']:
                                if main_list[main_data]['analytic_account_id'] in ColIndexes:
                                    mainlist_position[main_list[main_data]['analytic_account_id']]=ColIndexes[main_list[main_data]['analytic_account_id']]
                        for p,v in ColIndexes.items():
                            if p in mainlist_position:
                                for i in range(len(main_list)):
                                    if main_list[i]['analytic_account_id'] == p and main_list[i]['account_name'] == new_list[k]['account_name']:
                                        worksheet.write(row, v,abs(main_list[i]['balance']), floatstyle)
                            else:
                                worksheet.write(row, v, abs(00.0),rightfont)

                    if Projectwise == 'month':
                        mainlist_position={}
                        for main_data in range(len(main_list)):
                            if main_list[main_data]['account_name'] == new_list[k]['account_name']:
                                if main_list[main_data]['month'] in ColIndexes:
                                    mainlist_position[main_list[main_data]['month']]=ColIndexes[main_list[main_data]['month']]
                        for p,v in ColIndexes.items():
                            if p in mainlist_position:
                                for i in range(len(main_list)):
                                    if main_list[i]['month'] == p and main_list[i]['account_name'] == new_list[k]['account_name']:
                                        worksheet.write(row, v,abs(main_list[i]['balance']),floatstyle)
                            else:
                                worksheet.write(row, v, abs(00.0),rightfont)

                    if Projectwise == 'year':
                        mainlist_position={}
                        for main_data in range(len(main_list)):
                            if main_list[main_data]['account_name'] == new_list[k]['account_name']:
                                if main_list[main_data]['year'] in ColIndexes:
                                    mainlist_position[main_list[main_data]['year']]=ColIndexes[main_list[main_data]['year']]
                        for p,v in ColIndexes.items():
                            if p in mainlist_position:
                                for i in range(len(main_list)):
                                    if main_list[i]['year'] == p and main_list[i]['account_name'] == new_list[k]['account_name']:
                                        worksheet.write(row, v,abs(main_list[i]['balance']),floatstyle)
                            else:
                                worksheet.write(row, v, abs(00.0),rightfont)

                    if Projectwise == 'dimension' or Projectwise == 'month' or Projectwise == 'year':
                        for values in range(len(news_list)):
                            if news_list[values]['account_name'] == new_list[k]['account_name']:
                                total_column = news_list[values]['columns']
                                listd = [list(c.values())[0] for c in total_column]
                                third_income_list.append(listd)
                except:
                    pass
 
        incomeres  = []
        if Projectwise == 'dimension' or Projectwise == 'month' or Projectwise == 'year':
            if third_income_list:
                for j in range(0, len(third_income_list[0])): 
                    tmp = 0
                    for i in range(0, len(third_income_list)): 
                        tmp = tmp + third_income_list[i][j]
                    incomeres.append(tmp) 
     
        row+=1
        worksheet.write(row,1, 'Total Operating Income', style = mainheaders)
        worksheet.write(row,2, abs(OperatingIncome), style = mainheaderdata )
        worksheet.write(row,3, abs(100), style = mainheaderdata)
        col = 4
        if Projectwise == 'dimension'or Projectwise == 'month' or Projectwise == 'year':
            if incomeres:
                for j in range(len(incomeres)):
                    self.format_value(abs(incomeres[j]))
                    worksheet.write(row, col,abs(incomeres[j]), mainheaderdata)
                    col+=1
            else:
                for p,v in ColIndexes.items():
                    worksheet.write(row, col, abs(00.0), mainheaderdata)
                    col+=1
        row+=1
        worksheet.write(row, 0,'', style = mainheaders)
        worksheet.write(row,1, 'Cost of Revenue', style = mainheaders)
        worksheet.write(row, 2, '', style = mainheaderdata)
        worksheet.write(row,3, '', style = mainheaderdatas)
        for i in range(4,100):
          if i != col:
              worksheet.write(row, i,'',style = mainheaderdatas)
          elif i == col:
              break

        third_cost_list = []

        for k in range(0,len(new_list)):
            if new_list[k]['account_type'] == "Cost of Revenue":
                row+=1
                try:
                    percentage = round(((new_list[k]['balance'] * 100) / OperatingIncome),1)
                    worksheet.write(row, 0, new_list[k]['account_code'],alinedata)
                    worksheet.write(row, 1, new_list[k]['account_name'],alinedata)
                    worksheet.write(row, 2, abs(new_list[k]['balance']),floatstyle)
                    worksheet.write(row, 3, abs(percentage),linedata)
                    if Projectwise == 'dimension':
                        mainlist_position={}
                        for main_data in range(len(main_list)):
                            if main_list[main_data]['account_name'] == new_list[k]['account_name']:
                                if main_list[main_data]['analytic_account_id'] in ColIndexes:
                                    mainlist_position[main_list[main_data]['analytic_account_id']]=ColIndexes[main_list[main_data]['analytic_account_id']]
                        for p,v in ColIndexes.items():
                            if p in mainlist_position:
                                for i in range(len(main_list)):
                                    if main_list[i]['analytic_account_id'] == p and main_list[i]['account_name'] == new_list[k]['account_name']:
                                        worksheet.write(row, v,abs(main_list[i]['balance']),floatstyle)
                            else:
                                worksheet.write(row, v, abs(00.0),rightfont)

                    if Projectwise == 'month':
                        mainlist_position={}
                        for main_data in range(len(main_list)):
                            if main_list[main_data]['account_name'] == new_list[k]['account_name']:
                                if main_list[main_data]['month'] in ColIndexes:
                                    mainlist_position[main_list[main_data]['month']]=ColIndexes[main_list[main_data]['month']]
                        for p,v in ColIndexes.items():
                            if p in mainlist_position:
                                for i in range(len(main_list)):
                                    if main_list[i]['month'] == p and main_list[i]['account_name'] == new_list[k]['account_name']:
                                        worksheet.write(row, v,abs(main_list[i]['balance']),floatstyle)
                            else:
                                worksheet.write(row, v,abs(00.0),rightfont)

                    if Projectwise == 'year':
                        mainlist_position={}
                        for main_data in range(len(main_list)):
                            if main_list[main_data]['account_name'] == new_list[k]['account_name']:
                                if main_list[main_data]['year'] in ColIndexes:
                                    mainlist_position[main_list[main_data]['year']]=ColIndexes[main_list[main_data]['year']]
                        for p,v in ColIndexes.items():
                            if p in mainlist_position:
                                for i in range(len(main_list)):
                                    if main_list[i]['year'] == p and main_list[i]['account_name'] == new_list[k]['account_name']:
                                        worksheet.write(row, v,abs(main_list[i]['balance']),floatstyle)
                            else:
                                worksheet.write(row, v,abs(00.0),rightfont)

                    if Projectwise == 'dimension' or Projectwise == 'month' or Projectwise == 'year' :
                        for values in range(len(news_list)):
                            if news_list[values]['account_name'] == new_list[k]['account_name']:
                                total_column = news_list[values]['columns']
                                listd = [list(c.values())[0] for c in total_column]
                                third_cost_list.append(listd)
                except:
                    pass
        
        rescost  = []
        if Projectwise == 'dimension' or Projectwise == 'month' or Projectwise == 'year' :
            if third_cost_list:
                for j in range(0, len(third_cost_list[0])): 
                    tmp = 0
                    for i in range(0, len(third_cost_list)): 
                        tmp = tmp + third_cost_list[i][j]
                    rescost.append(tmp)
        row+=1
        worksheet.write(row,1, 'Total Cost of Revenue', style = mainheaders)
        worksheet.write(row,2, abs(CostOfRevenue), style = mainheaderdata )
        worksheet.write(row,3, round((abs(CorPercentage)),1), style = mainheaderdatas)
        col = 4
        if Projectwise == 'dimension' or Projectwise == 'month' or Projectwise == 'year' :
            if rescost:
                for j in range(len(rescost)):
                    worksheet.write(row, col,abs(rescost[j]), mainheaderdata)
                    col+=1
            else:
                for p,v in ColIndexes.items():
                    worksheet.write(row, col, abs(00.0),mainheaderdata)
                    col+=1

        row += 1
        res_list = []
        worksheet.write(row,1, 'Total Gross Profit', style = mainheaders)
        worksheet.write(row,2, GrossProfit, style = mainheaderdata )
        worksheet.write(row,3, round((GrossProfitPercentage),1), style = mainheaderdatas)
        if Projectwise == 'dimension' or Projectwise == 'month' or Projectwise == 'year':
            col = 4
            if rescost and incomeres :
                for i in range(0, len(rescost)):
                    res_list.append(incomeres[i] - rescost[i])

                for s in range(0, len(res_list)):
                    worksheet.write(row, col,res_list[s], mainheaderdata)
                    col+=1

            elif rescost:
                for i in range(0, len(rescost)):
                    
                    worksheet.write(row, col,rescost[s], mainheaderdata)
                    col+=1

            elif incomeres:
                for i in range(0, len(incomeres)):
                    worksheet.write(row, col,incomeres[i], mainheaderdata)
                    col+=1

        row+=1
        worksheet.write(row, 0,'', style = mainheaders)
        worksheet.write(row,1,'Other Income', style = mainheaders)
        worksheet.write(row,2,'', style = mainheaderdata)
        worksheet.write(row,3,'', style = mainheaderdatas)
        for i in range(4,100):
          if i != col:
              worksheet.write(row, i,'',style = mainheaderdata)
          elif i == col:
              break

        third_other_list = []

        for k in range(0,len(new_list)):
            if new_list[k]['account_type'] == "Other Income":
                row+=1
                try:
                    percentage = round(((new_list[k]['balance'] * 100) / OtherIncome),1)
                    worksheet.write(row, 0, new_list[k]['account_code'],alinedata)
                    worksheet.write(row, 1, new_list[k]['account_name'],alinedata)
                    worksheet.write(row, 2, abs(new_list[k]['balance']),floatstyle)
                    worksheet.write(row, 3, abs(percentage),floatstyle)
                    if Projectwise == 'dimension':
                        mainlist_position={}
                        for main_data in range(len(main_list)):
                            if main_list[main_data]['account_name'] == new_list[k]['account_name']:
                                if main_list[main_data]['analytic_account_id'] in ColIndexes:
                                    mainlist_position[main_list[main_data]['analytic_account_id']]=ColIndexes[main_list[main_data]['analytic_account_id']]
                        for p,v in ColIndexes.items():
                            if p in mainlist_position:
                                for i in range(len(main_list)):
                                    if main_list[i]['analytic_account_id'] == p and main_list[i]['account_name'] == new_list[k]['account_name']:
                                        worksheet.write(row, v,abs(main_list[i]['balance']),floatstyle)
                            else:
                                worksheet.write(row, v,abs(00.0),rightfont)
                    if Projectwise == 'month':
                        mainlist_position={}
                        for main_data in range(len(main_list)):
                            if main_list[main_data]['account_name'] == new_list[k]['account_name']:
                                if main_list[main_data]['month'] in ColIndexes:
                                    mainlist_position[main_list[main_data]['month']]=ColIndexes[main_list[main_data]['month']]
                        for p,v in ColIndexes.items():
                            if p in mainlist_position:
                                for i in range(len(main_list)):
                                    if main_list[i]['month'] == p and main_list[i]['account_name'] == new_list[k]['account_name']:
                                        worksheet.write(row, v,abs(main_list[i]['balance']),floatstyle)
                            else:
                                worksheet.write(row, v, abs(00.0),rightfont)
                    if Projectwise == 'year':
                        mainlist_position={}
                        for main_data in range(len(main_list)):
                            if main_list[main_data]['account_name'] == new_list[k]['account_name']:
                                if main_list[main_data]['year'] in ColIndexes:
                                    mainlist_position[main_list[main_data]['year']]=ColIndexes[main_list[main_data]['year']]
                        for p,v in ColIndexes.items():
                            if p in mainlist_position:
                                for i in range(len(main_list)):
                                    if main_list[i]['year'] == p and main_list[i]['account_name'] == new_list[k]['account_name']:
                                        worksheet.write(row, v, abs(main_list[i]['balance']),floatstyle)
                            else:
                                worksheet.write(row, v, abs(00.0),rightfont)

                    if Projectwise == 'dimension' or Projectwise == 'month' or Projectwise == 'year' :
                        for values in range(len(news_list)):
                            if news_list[values]['account_name'] == new_list[k]['account_name']:
                                total_column = news_list[values]['columns']
                                listd = [list(c.values())[0] for c in total_column]
                                third_other_list.append(listd)
                except:
                    pass
      
        resother  = []
        if Projectwise == 'dimension' or Projectwise == 'month' or Projectwise == 'year' :
            if third_other_list:
                for j in range(0, len(third_other_list[0])):
                    tmp = 0
                    for i in range(0, len(third_other_list)): 
                        tmp = tmp + third_other_list[i][j]
                    resother.append(tmp)

        row+=1
        worksheet.write(row,1, 'Total Other Income', style = mainheaders)
        worksheet.write(row,2, abs(OtherIncome), style = mainheaderdata )
        worksheet.write(row,3, round((00.0),1), style = mainheaderdatas)
        col = 4
        if Projectwise == 'dimension':
            if resother:
                for j in range(len(resother)):
                    worksheet.write(row, col,abs(resother[j]), mainheaderdata)
                    col+=1
            else:
                for p,v in ColIndexes.items():
                    worksheet.write(row, col,abs(00.0),mainheaderdata)
                    col+=1

        row +=1
        totalnetprofitloss = []
        worksheet.write(row,0, 'Total Income', style = mainheader)
        worksheet.write(row,1, '', style = mainheaderline)
        worksheet.write(row,2, abs(TotalIncome), style = mainheaderline )
        worksheet.write(row,3, round((TotalIncome * 100 / OperatingIncome),1), style = mainheaderline)
        if Projectwise == 'dimension' or Projectwise == 'month' or Projectwise == 'year':
            col = 4
            for i in range(0, len(finalincome_list)):
                worksheet.write(row, col,finalincome_list[i], mainheaderline)
                col+=1
        row+=1
        worksheet.write(row,0, 'Net Profit/Loss', style = mainheader)
        worksheet.write(row,1, '', style = mainheaderline)
        if NetProfit < 0:

            worksheet.write(row,2, NetProfit, style = mainheaderline)
        else :
            worksheet.write(row,2, '', style = mainheaderline)
        if NetProfit < 0:
            worksheet.write(row,3, round((NetProfit * 100 / OperatingIncome),1), style = mainheaderline)
        else:
            worksheet.write(row,3, '', style = mainheaderline)
        if Projectwise == 'dimension' or Projectwise == 'month' or Projectwise == 'year':
            col = 4
            for i in range(len(netbalance_list)):
                if netbalance_list[i] < 0:

                    worksheet.write(row, col, netbalance_list[i], style = mainheaderline)
                else:
                    worksheet.write(row,col, '', style = mainheaderline)

                col+=1
        row+=1
        worksheet.write(row, 0,'Expenses', style = mainheaderexpense)
        worksheet.write(row, 1,'', style = mainheaderexpense)
        worksheet.write(row, 2,'', style = mainheaderexpense)
        worksheet.write(row, 3,'', style = mainheaderexpense)
        for i in range(4,100):
          if i != col:
              worksheet.write(row, i,'',style = mainheaderexpense)
          elif i == col:
              break
        row +=1
        worksheet.write(row, 0,'', style = mainheaders)
        worksheet.write(row,1, 'Expenses',mainheaders)
        worksheet.write(row,2,'',mainheaderdata)
        worksheet.write(row,3,'',mainheaderdatas)
        for i in range(4,100):
          if i != col:
              worksheet.write(row, i,'',style = mainheaderdatas)
          elif i == col:
              break

        third_expense_list = []

        for k in range(0,len(new_list)):
            if  new_list[k]['account_type'] == "Expenses":
                row+=1
                try:
                    percentage = (new_list[k]['balance'] * 100 / OperatingIncome)
                    worksheet.write(row, 0, new_list[k]['account_code'],alinedata)
                    worksheet.write(row, 1, new_list[k]['account_name'],alinedata)
                    worksheet.write(row, 2, new_list[k]['balance'],floatstyle)
                    worksheet.write(row, 3, abs(percentage),floatstyle)
                    if Projectwise == 'dimension':
                        mainlist_position={}
                        for main_data in range(len(main_list)):
                            if main_list[main_data]['account_name'] == new_list[k]['account_name']:
                                if main_list[main_data]['analytic_account_id'] in ColIndexes:
                                    mainlist_position[main_list[main_data]['analytic_account_id']]=ColIndexes[main_list[main_data]['analytic_account_id']]
                        for p,v in ColIndexes.items():
                            if p in mainlist_position:
                                for i in range(len(main_list)):
                                    if main_list[i]['analytic_account_id'] == p and main_list[i]['account_name'] == new_list[k]['account_name']:
                                        worksheet.write(row, v, main_list[i]['balance'],floatstyle)
                            else:
                                worksheet.write(row, v, 00.0,rightfont)
                    if Projectwise == 'month':
                        mainlist_position={}
                        for main_data in range(len(main_list)):
                            if main_list[main_data]['account_name'] == new_list[k]['account_name']:
                                if main_list[main_data]['month'] in ColIndexes:
                                    mainlist_position[main_list[main_data]['month']]=ColIndexes[main_list[main_data]['month']]
                        for p,v in ColIndexes.items():
                            if p in mainlist_position:
                                for i in range(len(main_list)):
                                    if main_list[i]['month'] == p and main_list[i]['account_name'] == new_list[k]['account_name']:
                                        worksheet.write(row, v, main_list[i]['balance'],floatstyle)
                            else:
                                worksheet.write(row, v, 00.0, rightfont)
                    if Projectwise == 'year':
                        mainlist_position={}
                        for main_data in range(len(main_list)):
                            if main_list[main_data]['account_name'] == new_list[k]['account_name']:
                                if main_list[main_data]['year'] in ColIndexes:
                                    mainlist_position[main_list[main_data]['year']]=ColIndexes[main_list[main_data]['year']]
                        for p,v in ColIndexes.items():
                            if p in mainlist_position:
                                for i in range(len(main_list)):
                                    if main_list[i]['year'] == p and main_list[i]['account_name'] == new_list[k]['account_name']:
                                        worksheet.write(row, v, main_list[i]['balance'],floatstyle)
                            else:
                                worksheet.write(row, v, 00.0, rightfont)

                    if Projectwise == 'dimension' or Projectwise == 'month' or Projectwise == 'year':
                        for values in range(len(news_list)):
                            if news_list[values]['account_name'] == new_list[k]['account_name']:
                                total_column = news_list[values]['columns']
                                listd = [list(c.values())[0] for c in total_column]
                                third_expense_list.append(listd)
                except:
                    pass
       
        resexpense  = []
        if Projectwise == 'dimension' or Projectwise == 'month' or Projectwise == 'year':
            if third_expense_list:
                for j in range(0, len(third_expense_list[0])): 
                    tmp = 0
                    for i in range(0, len(third_expense_list)): 
                        tmp = tmp + third_expense_list[i][j]
                    resexpense.append(tmp)
        
        row+=1
        worksheet.write(row,1, 'Total Expenses', style = mainheaders)
        worksheet.write(row,2, Expenses, style = mainheaderdata )
        worksheet.write(row,3, round((ExpensePercentage),1), style = mainheaderdatas)
        col = 4
        if Projectwise == 'dimension' or Projectwise == 'month' or Projectwise == 'year':
            if resexpense:
                for j in range(len(resexpense)):
                  worksheet.write(row, col,resexpense[j], mainheaderdata)
                  col+=1
            else:
                for p,v in ColIndexes.items():
                    worksheet.write(row, col,abs(00.0), mainheaderdata)
                    col+=1

        row +=1
        worksheet.write(row, 0,'', style = mainheaders)      
        worksheet.write(row, 1, 'Depreciation', style = mainheaders)
        worksheet.write(row, 2, '', style = mainheaderdata)
        worksheet.write(row, 3, '', style = mainheaderdatas)
        for i in range(4,100):
          if i != col:
              worksheet.write(row, i,'',style = mainheaderdatas)
          elif i == col:
              break

        third_depriciation_list = []

        for k in range(0,len(new_list)):
            if new_list[k]['account_type'] == "Depreciation":
                row+=1
                try:
                    percentage = (new_list[k]['balance'] * 100 / OperatingIncome)
                    worksheet.write(row, 0, new_list[k]['account_code'],alinedata)
                    worksheet.write(row, 1, new_list[k]['account_name'],alinedata)
                    worksheet.write(row, 2, new_list[k]['balance'],floatstyle)
                    worksheet.write(row, 3, abs(percentage),floatstyle)
                    if Projectwise == 'dimension':
                        mainlist_position={}
                        for main_data in range(len(main_list)):
                            if main_list[main_data]['account_name'] == new_list[k]['account_name']:
                                if main_list[main_data]['analytic_account_id'] in ColIndexes:
                                    mainlist_position[main_list[main_data]['analytic_account_id']]=ColIndexes[main_list[main_data]['analytic_account_id']]
                        for p,v in ColIndexes.items():
                            if p in mainlist_position:
                                for i in range(len(main_list)):
                                    if main_list[i]['analytic_account_id'] == p and main_list[i]['account_name'] == new_list[k]['account_name']:
                                        worksheet.write(row, v, main_list[i]['balance'],floatstyle)
                            else:
                                worksheet.write(row, v, 00.0,rightfont)
                    if Projectwise == 'month':
                        mainlist_position={}
                        for main_data in range(len(main_list)):
                            if main_list[main_data]['account_name'] == new_list[k]['account_name']:
                                if main_list[main_data]['month'] in ColIndexes:
                                    mainlist_position[main_list[main_data]['month']]=ColIndexes[main_list[main_data]['month']]
                        for p,v in ColIndexes.items():
                            if p in mainlist_position:
                                for i in range(len(main_list)):
                                    if main_list[i]['month'] == p and main_list[i]['account_name'] == new_list[k]['account_name']:
                                        worksheet.write(row, v, main_list[i]['balance'],floatstyle)
                            else:
                                worksheet.write(row, v, 00.0, rightfont)
                    if Projectwise == 'year':
                        mainlist_position={}
                        for main_data in range(len(main_list)):
                            if main_list[main_data]['account_name'] == new_list[k]['account_name']:
                                if main_list[main_data]['year'] in ColIndexes:
                                    mainlist_position[main_list[main_data]['year']]=ColIndexes[main_list[main_data]['year']]
                        for p,v in ColIndexes.items():
                            if p in mainlist_position:
                                for i in range(len(main_list)):
                                    if main_list[i]['year'] == p and main_list[i]['account_name'] == new_list[k]['account_name']:
                                        worksheet.write(row, v, main_list[i]['balance'], floatstyle)
                            else:
                                worksheet.write(row, v, 00.0, rightfont)

                    if Projectwise == 'dimension' or Projectwise == 'month' or Projectwise == 'year':
                        for values in range(len(news_list)):
                            if news_list[values]['account_name'] == new_list[k]['account_name']:
                                total_column = news_list[values]['columns']
                                listd = [list(c.values())[0] for c in total_column]
                                third_depriciation_list.append(listd)

                except:
                    pass
       
        resdepriciation  = []
        if Projectwise == 'dimension' or Projectwise == 'month' or Projectwise == 'year':
            if third_depriciation_list: 
                for j in range(0, len(third_depriciation_list[0])):
                    tmp = 0
                    for i in range(0, len(third_depriciation_list)): 
                        tmp = tmp + third_depriciation_list[i][j]
                    resdepriciation.append(tmp)
            
        row+=1
        worksheet.write(row,1, 'Total Depreciation', style = mainheaders)
        worksheet.write(row,2, Depreciation, style = mainheaderdata)
        worksheet.write(row,3,round((DepPercentage),1), style = mainheaderdatas)
        col = 4
        if Projectwise == 'dimension' or Projectwise == 'month' or Projectwise == 'year':
            if resdepriciation:
                for j in range(len(resdepriciation)):
                    worksheet.write(row, col, resdepriciation[j], mainheaderdata)
                    col+=1
            else:
                for p,v in ColIndexes.items():
                    worksheet.write(row, col, 00.0, mainheaderdata)
                    col+=1
    
        row +=1
        worksheet.write(row,0, 'Total Expenses', style = mainheader)
        worksheet.write(row,1, '', style = mainheaderline)
        worksheet.write(row,2, TotalExpenses, style = mainheaderline)
        worksheet.write(row,3, round((TotalExpenses*100/OperatingIncome),1),  style = mainheaderline)
        if Projectwise == 'dimension' or Projectwise == 'month' or Projectwise == 'year':
            col = 4
            for i in range(0,len(finalexpense_list)):
                worksheet.write(row, col, finalexpense_list[i], style = mainheaderline)
                col+=1
           
        row +=1
        worksheet.write(row,0, 'Net Profit/Loss', style = netmainheaders)
        worksheet.write(row,1, '', style = netmainheader)
        if NetProfit > 0:
            worksheet.write(row,2, NetProfit, style = netmainheader)
        else :
            worksheet.write(row,2, '', style = netmainheader)
        if NetProfit > 0:
            worksheet.write(row,3, round((NetProfit*100/OperatingIncome),1) , style = netmainheader)
        else:
            worksheet.write(row,3, '', style = netmainheader)
        if Projectwise == 'dimension' or Projectwise == 'month' or Projectwise == 'year':
            col = 4
            for i in range(len(netbalance_list)):
                if netbalance_list[i] > 0:
                    worksheet.write(row, col, netbalance_list[i], style = netmainheader)
                else:
                    worksheet.write(row,col, '', style = netmainheader)

                col+=1
     
        row+=2
        buffer = io.BytesIO()
        workbook.save(buffer)
        export_id = self.env['profit.loss.excel'].create(
                        {'excel_file': base64.encodestring(buffer.getvalue()), 'file_name': filename})
        buffer.close()    
        return {
            'name': form_name,
            'view_mode': 'form',
            'res_id': export_id.id,
            'res_model': 'profit.loss.excel',
            'view_mode': 'form',
            'type': 'ir.actions.act_window',
            'target': 'new',
             }


class profit_loss_export_excel(models.TransientModel):
    _name= "profit.loss.excel"
    _description = "Profit And Loss Excel Report"

    excel_file = fields.Binary('Report for Profit And Loss')
    file_name = fields.Char('File', size=64)