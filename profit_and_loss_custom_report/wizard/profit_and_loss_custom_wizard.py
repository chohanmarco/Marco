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
# import numpy as np
# import pandas as pd
# from PIL import Image as PILImage
from odoo.tools.float_utils import float_round
from datetime import datetime
from dateutil.relativedelta import relativedelta

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


    def print_report(self):
      if self.date_from >= self.date_to:
          raise UserError(_("Start Date is greater than or equal to End Date."))
      datas = {'form': self.read()[0],
               'get_profit_loss': self.get_profit_and_loss_detail()
          }
      # print("datassss--",datas)
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
            for i in range(0,len(mainDict)):
                if (mainDict[i]['account_name'],mainDict[i]['analytic_account_id']) not in account_list:
                    main_list.append({
                                      'account_name':mainDict[i]['account_name'],
                                      'analytic_account_id':mainDict[i]['analytic_account_id'],
                                      'debit': mainDict[i]['account_debit'],
                                      'credit': mainDict[i]['account_credit'],
                                      'balance': mainDict[i]['account_debit'] - mainDict[i]['account_credit'] or 00.00,
                                      })
                    account_list.append((mainDict[i]['account_name'],mainDict[i]['analytic_account_id']))
                else:
                    first_list.append({
                                      'account_name':mainDict[i]['account_name'],
                                      'analytic_account_id':mainDict[i]['analytic_account_id'],
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
                                      'month': mainDict[i]['date'].strftime("%b %y")
                                      })
                    account_list.append((mainDict[i]['account_name'],mainDict[i]['date'].strftime("%b %y")))        
                else:
                    first_list.append({
                                      'account_name':mainDict[i]['account_name'],
                                      'debit': mainDict[i]['account_debit'],
                                      'credit': mainDict[i]['account_credit'],
                                      'balance': mainDict[i]['account_debit'] - mainDict[i]['account_credit'] or 00.00,
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
                                          'year': mainDict[i]['date'].strftime("%Y")
                                          })
                        account_list.append((mainDict[i]['account_name'],mainDict[i]['date'].strftime("%Y")))        
                else:
                    first_list.append({
                                      'account_name':mainDict[i]['account_name'],
                                      'debit': mainDict[i]['account_debit'],
                                      'credit': mainDict[i]['account_credit'],
                                      'balance': mainDict[i]['account_debit'] - mainDict[i]['account_credit'] or 00.00,
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
        linedata = xlwt.easyxf('borders: top_color black, bottom_color black, right_color black, left_color black,\
                              left thin, right thin, top thin, bottom thin; align: horiz right;')
        alinedata = xlwt.easyxf('borders: top_color black, bottom_color black, right_color black, left_color black,\
                              left thin, right thin, top thin, bottom thin; align: horiz left;')
        stylecolaccount = xlwt.easyxf('font: bold 1, colour white, height 200; \
                                      pattern: pattern solid, fore_colour dark_blue; \
                                      align: vert centre, horiz centre; \
                                      borders: top_color black, bottom_color black, right_color black, left_color black,\
                              left thin, right thin, top thin, bottom thin;')
        analytic_st_col = xlwt.easyxf('font: bold 1, colour black, height 200; \
                                    pattern: pattern solid, fore_colour gainsboro; \
                                    align: vert centre, horiz centre; \
                                    borders: top_color black, bottom_color black, right_color black, left_color black,\
                              left thin, right thin, top thin, bottom thin;')
        general = xlwt.easyxf('font: bold 1, colour black, height 210;')
        dateheader = xlwt.easyxf('font: bold 1, colour black, height 200;')
        maintotal = xlwt.easyxf('font: bold 1, colour black, height 200; \
                borders: top_color black, bottom_color black, right_color black, left_color black, \
        left thin, right thin, top thin, bottom thin;')
        finaltotalheader = xlwt.easyxf('pattern: fore_color white; font: bold 1, colour black; align: horiz right; \
        borders: top_color black, bottom_color black, right_color black, left_color black, \
        left thin, right thin, top thin, bottom thin;')
        rightfont = xlwt.easyxf('pattern: fore_color white; font: color dark_blue; align: horiz right; \
        borders: top_color black, bottom_color black, right_color black, left_color black, \
        left thin, right thin, top thin, bottom thin;')
        floatstyle = xlwt.easyxf("borders: top_color black, bottom_color black, right_color black, left_color black, \
        left thin, right thin, top thin, bottom thin;", "#,###.00")
        finaltotalheaderbold = xlwt.easyxf("pattern: fore_color white; font: bold 1, colour black; \
        borders: top_color black, bottom_color black, right_color black, left_color black, \
        left thin, right thin, top thin, bottom thin;", "#,###.00")
        accountnamestyle = xlwt.easyxf('font: bold 1, colour green, height 200;')
        mainheaders = xlwt.easyxf('pattern: fore_color white; font: bold 1, colour dark_blue; align: horiz left; borders: top_color black, bottom_color black, right_color black, left_color black,\
                              left thin, right thin, top thin, bottom thin;')
        mainheader = xlwt.easyxf('pattern: pattern solid, fore_colour gainsboro; \
                                 font: bold 1, colour dark_blue; align: horiz left; borders: top_color black, bottom_color black, right_color black, left_color black,\
                              left thin, right thin, top thin, bottom thin;')
        mainheaderline = xlwt.easyxf("pattern: pattern solid, fore_colour gainsboro; \
                                 font: bold 1, colour dark_blue; align: horiz right; borders: top_color black, bottom_color black, right_color black, left_color black,\
                              left thin, right thin, top thin, bottom thin;", "#,###.00")
        mainheaderdata = xlwt.easyxf("pattern: fore_color white; font: bold 1, colour dark_blue; align: horiz right; borders: top_color black, bottom_color black, right_color black, left_color black,left thin, right thin, top thin, bottom thin;", "#,###.00")
        mainheaderdatas = xlwt.easyxf("pattern: fore_color white; font: bold 1, colour dark_blue; align: horiz right; borders: top_color black, bottom_color black, right_color black, left_color black,left thin, right thin, top thin, bottom thin;")
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
        ColIndexes = { }
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
        totaldebit = 0.0
        totalcredit = 0.0
        totalbalance = 0.0
        totalanalytic = 0.0
        totalincomebalance = 0.0
        defaultpercentage = 100.00
        OpPercentage = 00.00
        OinPercentage = 00.00
        CorPercentage = 00.00
        ExPercentage = 00.00
        DepPercentage = 00.00
        OperatingIncome = 00.00
        OtherIncome = 00.00
        CostOfRevenue = 00.00
        ExpensesHeading = 00.00
        Depreciation = 00.00
        Expenses = 00.00
        IncomeHeading = 00.00
        GrossProfit = 00.00
        NetProfit = 00.00
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
                    OperatingIncome = OperatingIncome + new_list[k]['balance']
                elif new_list[k]['account_type'] == "Other Income":
                    OtherIncome = OtherIncome + new_list[k]['balance']
                elif new_list[k]['account_type'] == "Cost of Revenue":
                    CostOfRevenue += new_list[k]['balance']
                elif new_list[k]['account_type'] == "Depreciation":
                    Depreciation += new_list[k]['balance']
                elif new_list[k]['account_type'] == "Expenses":
                    Expenses += new_list[k]['balance']
                ExpensesHeading = Expenses + Depreciation
                GrossProfit = OperatingIncome + CostOfRevenue
                OpPercentage = round((OperatingIncome * 100 / GrossProfit),1)
                OinPercentage = round((OtherIncome * 100 / GrossProfit),1)
                CorPercentage = round((CostOfRevenue * 100 / GrossProfit),1)
                ExPercentage = round((Expenses * 100 / GrossProfit),1)
                DepPercentage = round((Depreciation * 100 / GrossProfit),1)
                NetProfit = GrossProfit + ExpensesHeading + OtherIncome
            except :
                pass

        worksheet.write_merge(row, row, 0,colc, 'Income', style = mainheader)
        row +=1
        worksheet.write(row, 0,'', style = mainheaders)
        worksheet.write(row, 1,'Gross Profit', style = mainheaders)
        worksheet.write(row, 2,GrossProfit, style = mainheaderdata)
        worksheet.write(row, 3,str(defaultpercentage), style = mainheaderdata)
        for i in range(4,100):
          if i != col:
              worksheet.write(row, i,0.0,style = mainheaderdatas)
          elif i == col:
              break
        row +=1
        worksheet.write(row, 0,'', style = mainheaders)
        worksheet.write(row,1, 'Operating Income', style = mainheaders)
        worksheet.write(row,2, OperatingIncome, style = mainheaderdata )
        worksheet.write(row,3, OpPercentage, style = mainheaderdatas)
        for i in range(4,100):
          if i != col:
              worksheet.write(row, i,0.0,style = mainheaderdatas)
          elif i == col:
              break
        for k in range(0,len(new_list)):
            if  new_list[k]['account_type'] == "Income":
                row+=1
                try:
                    percentage = round(((new_list[k]['balance'] * 100) / GrossProfit),1)
                    worksheet.write(row, 0, new_list[k]['account_code'],alinedata)
                    worksheet.write(row, 1, new_list[k]['account_name'],alinedata)
                    worksheet.write(row, 2, new_list[k]['balance'],floatstyle)
                    worksheet.write(row, 3, percentage,linedata)         
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
                                worksheet.write(row, v, 0.0,rightfont)
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
                                worksheet.write(row, v, 0.0,rightfont)
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
                                worksheet.write(row, v, 0.0,rightfont)
                except:
                    pass

        row+=1
        worksheet.write(row, 0,'', style = mainheaders)
        worksheet.write(row,1, 'Cost of Revenue', style = mainheaders)
        worksheet.write(row, 2, CostOfRevenue, style = mainheaderdata)
        worksheet.write(row,3, CorPercentage, style = mainheaderdatas)
        for i in range(4,100):
          if i != col:
              worksheet.write(row, i,0.0,style = mainheaderdatas)
          elif i == col:
              break
        for k in range(0,len(new_list)):
            if new_list[k]['account_type'] == "Cost of Revenue":
                row+=1
                try:
                    percentage = round(((new_list[k]['balance'] * 100) / GrossProfit),1)
                    worksheet.write(row, 0, new_list[k]['account_code'],alinedata)
                    worksheet.write(row, 1, new_list[k]['account_name'],alinedata)
                    worksheet.write(row, 2, new_list[k]['balance'],floatstyle)
                    worksheet.write(row, 3, percentage,linedata)
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
                                worksheet.write(row, v,0.0,rightfont)

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
                                worksheet.write(row, v, 0.0,rightfont)

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
                                worksheet.write(row, v, 0.0,rightfont)
                except:
                    pass

        row += 1
        worksheet.write(row, 0,'', style = mainheaders)
        worksheet.write(row,1,'Other Income', style = mainheaders)
        worksheet.write(row,2,OtherIncome, style = mainheaderdata)
        worksheet.write(row,3,OinPercentage, style = mainheaderdatas)
        for i in range(4,100):
          if i != col:
              worksheet.write(row, i,0.0,style = mainheaderdatas)
          elif i == col:
              break
        for k in range(0,len(new_list)):
            if new_list[k]['account_type'] == "Other Income":
                row+=1
                try:
                    percentage = round(((new_list[k]['balance'] * 100) / OtherIncome),1)
                    worksheet.write(row, 0, new_list[k]['account_code'],alinedata)
                    worksheet.write(row, 1, new_list[k]['account_name'],alinedata)
                    worksheet.write(row, 2, new_list[k]['balance'],floatstyle)
                    worksheet.write(row, 3, percentage,linedata)
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
                                worksheet.write(row, v, 0.0,rightfont)
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
                                worksheet.write(row, v, 0.0,rightfont)
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
                                worksheet.write(row, v, 0.0,rightfont)
                except:
                    pass

        row +=1
        worksheet.write_merge(row, row, 0,colc, 'Expenses', style = mainheader)
        worksheet.write(row,2,ExpensesHeading, mainheaderdata)
        row +=1
        worksheet.write(row, 0,'', style = mainheaders)
        worksheet.write(row,1, 'Expenses',mainheaders)
        worksheet.write(row,2,Expenses,mainheaderdata)
        worksheet.write(row,3,ExPercentage,mainheaderdatas)
        for i in range(4,100):
          if i != col:
              worksheet.write(row, i,0.0,style = mainheaderdatas)
          elif i == col:
              break
    
        for k in range(0,len(new_list)):
            if  new_list[k]['account_type'] == "Expenses":
                row+=1
                try:
                    percentage = round(((new_list[k]['balance'] * 100) / GrossProfit),1)
                    worksheet.write(row, 0, new_list[k]['account_code'],alinedata)
                    worksheet.write(row, 1, new_list[k]['account_name'],alinedata)
                    worksheet.write(row, 2, new_list[k]['balance'],floatstyle)
                    worksheet.write(row, 3, percentage,linedata)
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
                                worksheet.write(row, v, 0.0,rightfont)
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
                                worksheet.write(row, v, 0.0,rightfont)
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
                                worksheet.write(row, v, 0.0,rightfont)
                except:
                    pass
        row +=1
        worksheet.write(row, 0,'', style = mainheaders)      
        worksheet.write(row, 1, 'Depreciation', style = mainheaders)
        worksheet.write(row, 2, Depreciation, style = mainheaderdata)
        worksheet.write(row, 3, DepPercentage, style = mainheaderdatas)
        for i in range(4,100):
          if i != col:
              worksheet.write(row, i,0.0,style = mainheaderdatas)
          elif i == col:
              break
        for k in range(0,len(new_list)):
            if new_list[k]['account_type'] == "Depreciation":
                row+=1
                try:
                    percentage = round(((new_list[k]['balance'] * 100) / GrossProfit),1)
                    worksheet.write(row, 0, new_list[k]['account_code'],alinedata)
                    worksheet.write(row, 1, new_list[k]['account_name'],alinedata)
                    worksheet.write(row, 2, new_list[k]['balance'],floatstyle)
                    worksheet.write(row, 3, percentage,linedata)
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
                                worksheet.write(row, v, 0.0,rightfont)
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
                                worksheet.write(row, v, 0.0,rightfont)
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
                                worksheet.write(row, v,0.0,rightfont)
                except:
                    pass
        row +=1
        worksheet.write(row,0, 'Net Profit', style = mainheader)
        worksheet.write(row,1, '', style = mainheaderline)
        worksheet.write(row,2, NetProfit, style = mainheaderline)
        worksheet.write(row,3, '', style = mainheaderline)
        for i in range(4,100):
          if i != col:
              worksheet.write(row, i,0.0,style = mainheaderdatas)
          elif i == col:
              break
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