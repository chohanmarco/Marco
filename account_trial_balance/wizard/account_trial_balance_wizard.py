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
import numpy as np
import pandas as pd
from PIL import Image as PILImage
from odoo.tools.float_utils import float_round
from dateutil.rrule import rrule, MONTHLY
import calendar
from datetime import timedelta
from dateutil.relativedelta import relativedelta
from odoo.tools.misc import formatLang


class AccountTrialBalanceReport(models.TransientModel):
    _name = 'account.trial.balance.report'
    _description = "Account Trial Balance Report"

    date_from = fields.Date(string="From Date")
    date_to = fields.Date(string="To Date")
    account_ids = fields.Many2many('account.account', string='Accounts')
    account_without_transaction = fields.Boolean(string= 'Show Accounts without transactions', default=False)
    account_zero_closing_balance = fields.Boolean(string= 'Show Accounts with zero closing balance', default=False)
    dimension_wise_project = fields.Selection([('none','None'),('month','Month Wise'),('dimension', 'Dimension')],
                                              string='Dimension',
                                              default='none')
    projectwise = fields.Selection([('project', 'Project')],string='Project',default='project')
    detail_report = fields.Boolean(string= 'Show Detail Report(Accounting)', default=False)
    show_dr_cr_separately = fields.Boolean(string= 'Show Opening Dr/Cr Separately', default=False)
    analytic_account_ids = fields.Many2many('account.analytic.account', string='Analytic Accounts')


    def action_redirect_to_aml_view(self):
        action = self.env.ref('account_report_customization.action_account_moves_all_filter_with_report_tree').read()[0]
        dateFrom = self.date_from
        dateTo = self.date_to
        AllAccounts = self.account_ids
        FilteredAccountIds = AllAccounts.filtered(lambda a: a.temp_for_report)
        AccountIds = FilteredAccountIds.ids
        if not AccountIds:
            AccountIds = AllAccounts.ids
        Status = ['posted']
        MoveLines = []
        if AccountIds:
            self.env.cr.execute("""
                SELECT aml.id
                FROM account_move_line aml
                LEFT JOIN account_move am ON (am.id=aml.move_id)
                WHERE (aml.date >= %s) AND
                    (aml.date <= %s) AND
                    (aml.account_id in %s) AND
                    (am.state in %s) ORDER BY aml.date""",
                (str(dateFrom) + ' 00:00:00', str(dateTo) + ' 23:59:59', tuple(AccountIds), tuple(Status),))
            MoveLines = [x[0] for x in self.env.cr.fetchall()]
        action['context'] = {'create': False}
        action['domain'] = [('id', 'in', MoveLines)]
        return action

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

 
    def get_trial_balance_detail(self):
        """ Details For PDF Report """
        new_lines = []
        AccountGroupObj = self.env['account.group']
        GroupIds = AccountGroupObj.search([])
        dateFrom = self.date_from
        dateTo = self.date_to
        dates = {}
        AllAnalyticAccounts = self.analytic_account_ids
        FilteredAnalyticAccountIds = AllAnalyticAccounts.filtered(lambda a: a.temp_analytic_report)
        AnalyticAccountIds = FilteredAnalyticAccountIds
        AnalyticNames = []
        AnalyticIds = []
        if not AnalyticAccountIds:
            AnalyticAccountIds = AllAnalyticAccounts
        if self.dimension_wise_project == 'dimension':
            AnalyticIds = [analytic_account.id for analytic_account in AnalyticAccountIds]
            AnalyticNames = [analytic_account.name for analytic_account in AnalyticAccountIds]
        if self.dimension_wise_project == 'month':
               dates = {'date_from': dateFrom.strftime('%Y-%m-%d'),'date_to':dateTo.strftime('%Y-%m-%d')}
        CompanyImage = self.env.company.logo
        group_list = []
        option_dict = {}
        string = dateFrom.strftime('%Y')
        Status = ['posted']
        initial_balances = [True]  
        accounts = []
        accounts_results = []
        for group_ids in GroupIds:
            group_list.append(group_ids.name)
        queries = []
        option_dict.update({
                            'unfolded_lines':group_list,
                            'date':{'string': string,'mode':'range','date_from': dateFrom.strftime('%Y-%m-%d'),'date_to':dateTo.strftime('%Y-%m-%d')},
                            'analytic_accounts': AnalyticIds,
                            'analytic_accounts_name':AnalyticNames,
                            'month_wise_dates': dates,
                            })

        option_list = [option_dict]

        query, params = self.get_all_queries(option_list)
        groupby_accounts = {}
        groupby_companies = {}
        groupby_taxes = {}

        self._cr.execute(query, params)
        for res in self.env.cr.dictfetchall():

            if res['groupby'] is None:
                continue

            i = res['period_number']
            key = res['key']
            if key == 'sum':
                groupby_accounts.setdefault(res['groupby'], [{} for n in range(len(option_list))])
                groupby_accounts[res['groupby']][i][key] = res
            elif key == 'initial_balance':
                groupby_accounts.setdefault(res['groupby'], [{} for n in range(len(option_list))])
                groupby_accounts[res['groupby']][i][key] = res
            elif key == 'unaffected_earnings':
                groupby_companies.setdefault(res['groupby'], [{} for n in range(len(option_list))])
                groupby_companies[res['groupby']][i] = res
            elif key == 'dimensionsum':
                groupby_accounts.setdefault(res['groupby'], [{} for n in range(len(option_list))])
                groupby_accounts[res['groupby']][i][key] = res

        if groupby_companies:
            unaffected_earnings_type = self.env.ref('account.data_unaffected_earnings')
            candidates_accounts = self.env['account.account'].search([
                ('user_type_id', '=', unaffected_earnings_type.id), ('company_id', 'in', list(groupby_companies.keys()))
            ])
            for account in candidates_accounts:
                company_unaffected_earnings = groupby_companies.get(account.company_id.id)
                if not company_unaffected_earnings:
                    continue
                for i in range(len(option_list)):
                    unaffected_earnings = company_unaffected_earnings[i]
                    groupby_accounts.setdefault(account.id, [{} for i in range(len(option_list))])
                    groupby_accounts[account.id][i]['unaffected_earnings'] = unaffected_earnings
                del groupby_companies[account.company_id.id]

        AccountIds = []
        AllAccounts = self.account_ids
        FilteredAccountIds = AllAccounts.filtered(lambda a: a.temp_for_report)
        for ids in FilteredAccountIds.ids:
            ac = [i for i in list(groupby_accounts.keys())]
            if ids in ac:
                AccountIds.append(ids)
        if not AccountIds:
            for i in AllAccounts:
                if i.id in list(groupby_accounts.keys()):
                    AccountIds = [i for i in list(groupby_accounts.keys())]
        if groupby_accounts:
            accounts = self.env['account.account'].browse(AccountIds)
        else:
            accounts = []

        if groupby_accounts:
            accounts_results = [(account, groupby_accounts[account.id]) for account in accounts]
        lines = []
        totals = [0.0] * (2 * (len(option_list) + 2))
        for account, periods_results in accounts_results:
            sums = []
            account_balance = 0.0
            for i, period_values in enumerate(reversed(periods_results)):
                account_sum = period_values.get('sum', {})
                account_un_earn = period_values.get('unaffected_earnings', {})
                account_init_bal = period_values.get('initial_balance', {})

                if i == 0:
                    initial_balance = account_init_bal.get('balance', 0.0) + account_un_earn.get('balance', 0.0)
                    sums += [
                        initial_balance > 0 and initial_balance or 0.0,
                        initial_balance < 0 and -initial_balance or 0.0,
                    ]
                    account_balance += initial_balance

                # Append the debit/credit columns.
                sums += [
                    account_sum.get('debit', 0.0) - account_init_bal.get('debit', 0.0),
                    account_sum.get('credit', 0.0) - account_init_bal.get('credit', 0.0),
                ]
                account_balance += sums[-2] - sums[-1]

            # Append the totals.
            sums += [
                account_balance > 0 and account_balance or 0.0,
                account_balance < 0 and -account_balance or 0.0,
            ]
            # account.account report line.
            columns = []
            for i, value in enumerate(sums):
                # Update totals.
                totals[i] += value

                # Create columns.
                columns.append({'name': value, 'class': 'number', 'no_format_name': value})

            name = account.name_get()[0][1]
            code = name.split()[0]
            if len(name) > 40 and not self._context.get('print_mode'):
                name = name[:40]+'...'

            lines.append({
                'id': account.id,
                'code': code,
                'name': name,
                'title_hover': name,
                'columns': columns,
                'unfoldable': False,
                'caret_options': 'account.account',
            })

        # Total report line.
        lines.append({
             'id': 'grouped_accounts_total',
             'code': 'group_code',
             'name': _('Total'),
             'class': 'total',
             'columns': [{'name': total, 'class': 'number'} for total in totals],
             'level': 1,
        })

        accounts_hierarchy = {}
        no_group_lines = []
        for line in lines + [None]:
           
            is_grouped_by_account = line and (line.get('caret_options') == 'account.account' or line.get('account_id'))
            if not is_grouped_by_account or not line:
                no_group_hierarchy = {}
                for no_group_line in no_group_lines:
                    codes = [('root', str(line.get('parent_id')) or 'root') if line else 'root', (self.LEAST_SORT_PRIO, _('(No Group)'))]
                    if not accounts_hierarchy:
                        account = self.get_account(no_group_line.get('account_id', no_group_line.get('id')))
                        codes = [('root', line and str(line.get('parent_id')) or 'root')] + self.get_account_codes(account)
                    self.add_line_to_hierarchy(no_group_line, codes, no_group_hierarchy, line and line.get('level') or 0 + 1)
                no_group_lines = []
               
                self.deep_merge_dict(no_group_hierarchy, accounts_hierarchy)
                # Merge the newly created hierarchy with existing lines.
                if accounts_hierarchy:
                   
                    new_lines += self.get_hierarchy_lines(accounts_hierarchy)[1]
                   
                    accounts_hierarchy = {}

                if line:
                    new_lines.append(line)

                continue

            # Exclude lines having no group.
            account = self.get_account(line.get('account_id', line.get('id')))

            if not account.group_id.id:
                
                no_group_lines.append(line)
                continue
            codes = [('root', str(line.get('parent_id')) or 'root')] + self.get_account_codes(account)
            self.add_line_to_hierarchy(line, codes, accounts_hierarchy,line.get('level', 0) + 1)

        net_balance = 0.0
        for i in range(len(new_lines)):
            acc_balance = new_lines[i]['columns']
            listd = [list(c.values())[0] for c in acc_balance]
            first = listd[4]
            second = listd[5]
            net_balance = first - second
            acc_balance .append({'name':net_balance, 'class': 'number', 'no_format_name':net_balance})

        for i in range(len(new_lines)):
          if not self.show_dr_cr_separately:
            new_lines[i]['columns']=new_lines[i]['columns'][2:]

        if not self.account_without_transaction :
            for val in new_lines:
                v = 0.0 
                acc_balance = val['columns']
                for k in acc_balance:
                    v +=list(k.values())[0]

                if v == 0.0 or v == 0 :
                    new_lines.remove(val)

        return new_lines

    @api.model
    def default_get(self, fields):
        vals = super(AccountTrialBalanceReport, self).default_get(fields)
        ac_ids = self.env['account.account'].search([])
        analytic_ids = self.env['account.analytic.account'].search([])
        self.env.cr.execute('update account_account set temp_for_report=False')
        self.env.cr.execute('update account_analytic_account set temp_analytic_report=False')
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

    def get_account(self,id):
        accounts_cache = {}
        if id not in accounts_cache:
            accounts_cache[id] = self.env['account.account'].browse(id)
        return accounts_cache[id]

    def get_account_codes(self, account):
        # A code is tuple(sort priority, actual code)
        codes = []
        # if not account:
        #     continue
        if account.group_id:
            group = account.group_id
            while group:
                code = '%s %s' % (group.code_prefix or '', group.name)
                codes.append((self.MOST_SORT_PRIO, code))
                group = group.parent_id
        else:
            # Limit to 3 levels.
            code = account.code[:3]
            while code:
                codes.append((self.MOST_SORT_PRIO, code))
                code = code[:-1]
        return list(reversed(codes))

    MOST_SORT_PRIO = 0
    LEAST_SORT_PRIO = 99

    def merge_columns(self,columns):
        return [('n/a' if any(i != '' for i in x) else '') if any(isinstance(i, str) for i in x) else sum(x) for x in zip(*columns)]
            
    def get_hierarchy_lines(self,values, depth=1):
        
            lines = []
            sum_sum_columns = []
            # unfold_all = self.env.context.get('print_mode') and len(options.get('unfolded_lines')) == 0
            for base_line in values.get('lines', []):
                lines.append(base_line)
                sum_sum_columns.append([c.get('no_format_name', c['name']) for c in base_line['columns']])
            
            # For the last iteration, there might not be the children key (see add_line_to_hierarchy)
            for key in sorted(values.get('children', {}).keys()):
                sum_columns, sub_lines = self.get_hierarchy_lines(values['children'][key], depth=values['depth'])
                
                id = 'hierarchy_' + key[1]
                acc_code = key[1].split(" ")[0]
                header_line = {
                    'id': id,
                    'code':acc_code,
                    'name': key[1] if len(key[1]) < 60 else key[1][:60] + '...',  # second member of the tuple
                    'title_hover': key[1],
                    'unfoldable': True,
                    'level': values['depth'],
                    'parent_id': values['parent_id'],
                    'columns': [{'name': c if not isinstance(c, str) else c} for c in sum_columns],
                }

                if key[0] == self.LEAST_SORT_PRIO:
                    header_line['style'] = 'font-style:italic;'
                lines += [header_line] + sub_lines
                sum_sum_columns.append(sum_columns)
            return self.merge_columns(sum_sum_columns), lines


    def add_line_to_hierarchy(self,line, codes, level_dict, depth=None):
       
        if not codes:
            return
        if not depth:
            depth = line.get('level', 1)
        level_dict.setdefault('depth', depth)
        
        level_dict.setdefault('parent_id', 'hierarchy_' + codes[0][1] if codes[0][0] != 'root' else codes[0][1])
       
        level_dict.setdefault('children', {})
        
        code = codes[1]
        codes = codes[1:]
        level_dict['children'].setdefault(code, {})
       
        if len(codes) > 1:
            self.add_line_to_hierarchy(line, codes, level_dict['children'][code], depth=depth + 1)
            
        else:
            level_dict['children'][code].setdefault('lines', [])
           
            level_dict['children'][code]['lines'].append(line)
            
            line['level'] = depth + 1
            
            for l in level_dict['children'][code]['lines']:
                l['parent_id'] = 'hierarchy_' + code[1]
               

    def deep_merge_dict(self, source, destination):
        for key, value in source.items():
            if isinstance(value, dict):
                # get node or create one
                node = destination.setdefault(key, {})
                self.deep_merge_dict(value, node)
            else:
                destination[key] = value
        return destination

    def trial_balance_export_excel(self):
        """
        This methods make list of dict to Export in Dailybook Excel
        """
        new_lines = []
        AccountGroupObj = self.env['account.group']
        GroupIds = AccountGroupObj.search([])
        dateFrom = self.date_from
        dateTo = self.date_to
        dates = {}
        AllAnalyticAccounts = self.analytic_account_ids
        FilteredAnalyticAccountIds = AllAnalyticAccounts.filtered(lambda a: a.temp_analytic_report)
        AnalyticAccountIds = FilteredAnalyticAccountIds
        AnalyticNames = []
        AnalyticIds = []
        if not AnalyticAccountIds:
            AnalyticAccountIds = AllAnalyticAccounts
        if self.dimension_wise_project == 'dimension':
            AnalyticIds = [analytic_account.id for analytic_account in AnalyticAccountIds]
            AnalyticNames = [analytic_account.name for analytic_account in AnalyticAccountIds]
        if self.dimension_wise_project == 'month':
               dates = {'date_from': dateFrom.strftime('%Y-%m-%d'),'date_to':dateTo.strftime('%Y-%m-%d')}
        CompanyImage = self.env.company.logo
        group_list = []
        option_dict = {}
        string = dateFrom.strftime('%Y')
        Status = ['posted']
        initial_balances = [True]  
        accounts = []
        for group_ids in GroupIds:
            group_list.append(group_ids.name)
        queries = []
        option_dict.update({
                            'unfolded_lines':group_list,
                            'date':{'string': string,'mode':'range','date_from': dateFrom.strftime('%Y-%m-%d'),'date_to':dateTo.strftime('%Y-%m-%d')},
                            'analytic_accounts': AnalyticIds,
                            'analytic_accounts_name':AnalyticNames,
                            'month_wise_dates': dates,
                            })

        option_list = [option_dict]

        query, params = self.get_all_queries(option_list)
        groupby_accounts = {}
        groupby_companies = {}
        groupby_taxes = {}

        self._cr.execute(query, params)
        for res in self.env.cr.dictfetchall():

            if res['groupby'] is None:
                continue

            i = res['period_number']
            key = res['key']
            if key == 'sum':
                groupby_accounts.setdefault(res['groupby'], [{} for n in range(len(option_list))])
                groupby_accounts[res['groupby']][i][key] = res
            elif key == 'initial_balance':
                groupby_accounts.setdefault(res['groupby'], [{} for n in range(len(option_list))])
                groupby_accounts[res['groupby']][i][key] = res
            elif key == 'unaffected_earnings':
                groupby_companies.setdefault(res['groupby'], [{} for n in range(len(option_list))])
                groupby_companies[res['groupby']][i] = res
            elif key == 'dimensionsum':
                groupby_accounts.setdefault(res['groupby'], [{} for n in range(len(option_list))])
                groupby_accounts[res['groupby']][i][key] = res

        if groupby_companies:
            unaffected_earnings_type = self.env.ref('account.data_unaffected_earnings')
            candidates_accounts = self.env['account.account'].search([
                ('user_type_id', '=', unaffected_earnings_type.id), ('company_id', 'in', list(groupby_companies.keys()))
            ])
            for account in candidates_accounts:
                company_unaffected_earnings = groupby_companies.get(account.company_id.id)
                if not company_unaffected_earnings:
                    continue
                for i in range(len(option_list)):
                    unaffected_earnings = company_unaffected_earnings[i]
                    groupby_accounts.setdefault(account.id, [{} for i in range(len(option_list))])
                    groupby_accounts[account.id][i]['unaffected_earnings'] = unaffected_earnings
                del groupby_companies[account.company_id.id]

        AccountIds = []
        AllAccounts = self.account_ids
        FilteredAccountIds = AllAccounts.filtered(lambda a: a.temp_for_report)
        for ids in FilteredAccountIds.ids:
            ac = [i for i in list(groupby_accounts.keys())]
            if ids in ac:
                AccountIds.append(ids)
        if not AccountIds:
            for i in AllAccounts:
                if i.id in list(groupby_accounts.keys()):
                    AccountIds = [i for i in list(groupby_accounts.keys())]
        if groupby_accounts:
            accounts = self.env['account.account'].browse(AccountIds)
        else:
            accounts = []

        if groupby_accounts:
            accounts_results = [(account, groupby_accounts[account.id]) for account in accounts]
        lines = []
        totals = [0.0] * (2 * (len(option_list) + 2))
        for account, periods_results in accounts_results:
            sums = []
            account_balance = 0.0
            for i, period_values in enumerate(reversed(periods_results)):
                account_sum = period_values.get('sum', {})
                account_un_earn = period_values.get('unaffected_earnings', {})
                account_init_bal = period_values.get('initial_balance', {})

                if i == 0:
                    initial_balance = account_init_bal.get('balance', 0.0) + account_un_earn.get('balance', 0.0)
                    sums += [
                        initial_balance > 0 and initial_balance or 0.0,
                        initial_balance < 0 and -initial_balance or 0.0,
                    ]
                    account_balance += initial_balance

                # Append the debit/credit columns.
                sums += [
                    account_sum.get('debit', 0.0) - account_init_bal.get('debit', 0.0),
                    account_sum.get('credit', 0.0) - account_init_bal.get('credit', 0.0),
                ]
                account_balance += sums[-2] - sums[-1]

            # Append the totals.
            sums += [
                account_balance > 0 and account_balance or 0.0,
                account_balance < 0 and -account_balance or 0.0,
            ]
            # account.account report line.
            columns = []
            for i, value in enumerate(sums):
                # Update totals.
                totals[i] += value

                # Create columns.
                columns.append({'name': value, 'class': 'number', 'no_format_name': value})

            name = account.name_get()[0][1]
            code = name.split()[0]
            if len(name) > 40 and not self._context.get('print_mode'):
                name = name[:40]+'...'

            lines.append({
                'id': account.id,
                'code': code,
                'name': name,
                'title_hover': name,
                'columns': columns,
                'unfoldable': False,
                'caret_options': 'account.account',
            })

        # Total report line.
        lines.append({
             'id': 'grouped_accounts_total',
             'code': 'group_code',
             'name': _('Total'),
             'class': 'total',
             'columns': [{'name': total, 'class': 'number'} for total in totals],
             'level': 1,
        })

        accounts_hierarchy = {}
        no_group_lines = []
        for line in lines + [None]:
           
            is_grouped_by_account = line and (line.get('caret_options') == 'account.account' or line.get('account_id'))
            if not is_grouped_by_account or not line:
                no_group_hierarchy = {}
                for no_group_line in no_group_lines:
                    codes = [('root', str(line.get('parent_id')) or 'root') if line else 'root', (self.LEAST_SORT_PRIO, _('(No Group)'))]
                    if not accounts_hierarchy:
                        account = self.get_account(no_group_line.get('account_id', no_group_line.get('id')))
                        codes = [('root', line and str(line.get('parent_id')) or 'root')] + self.get_account_codes(account)
                    self.add_line_to_hierarchy(no_group_line, codes, no_group_hierarchy, line and line.get('level') or 0 + 1)
                no_group_lines = []
               
                self.deep_merge_dict(no_group_hierarchy, accounts_hierarchy)
                # Merge the newly created hierarchy with existing lines.
                if accounts_hierarchy:
                   
                    new_lines += self.get_hierarchy_lines(accounts_hierarchy)[1]
                   
                    accounts_hierarchy = {}

                if line:
                    new_lines.append(line)

                continue

            # Exclude lines having no group.
            account = self.get_account(line.get('account_id', line.get('id')))

            if not account.group_id.id:
                
                no_group_lines.append(line)
                continue
            codes = [('root', str(line.get('parent_id')) or 'root')] + self.get_account_codes(account)
            self.add_line_to_hierarchy(line, codes, accounts_hierarchy,line.get('level', 0) + 1)
        
        import base64
        filename = 'Trial Balance.xls'
        start = dateFrom.strftime('%d %b, %Y')
        end = dateTo.strftime('%d %b, %Y')
        form_name = 'Trial Balance Between ' + str(start) + ' to ' + str(end)
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

        worksheet = workbook.add_sheet("Trial Balance", cell_overwrite_ok=True)
        worksheet.show_grid = False

        styleheader = xlwt.easyxf('font: bold 1, colour black, height 300;')
        
        
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
       
        mainheaderdata = xlwt.easyxf('borders: top_color black, bottom_color black, right_color black, left_color black,\
                              left thin, right thin, top thin, bottom thin; align: horiz left;',)

        mainheader = xlwt.easyxf('pattern: pattern solid, fore_colour gainsboro; \
                                 font: bold 1, colour dark_blue; align: horiz left; borders: top_color black, bottom_color black, right_color black, left_color black,\
                              left thin, right thin, top thin, bottom thin;')

        mainheaders = xlwt.easyxf('pattern: fore_color white; font: bold 1, colour dark_blue; align: horiz left; borders: top_color black, bottom_color black, right_color black, left_color black,\
                              left thin, right thin, top thin, bottom thin;',)

        zero_col = worksheet.col(0)
        zero_col.width = 236 * 20
        first_col = worksheet.col(1)
        first_col.width = 236 * 40
        second_col = worksheet.col(2)
        second_col.width = 236 * 20
        third_col = worksheet.col(3)
        third_col.width = 236 * 20
        fourth_col = worksheet.col(4)
        fourth_col.width = 236 * 20
        fifth_col = worksheet.col(5)
        fifth_col.width = 236 * 20
        sixth_col = worksheet.col(6)
        sixth_col.width = 236 * 20
        seventh_col = worksheet.col(7)
        seventh_col.width = 236 * 20
        eigth_col = worksheet.col(8)
        eigth_col.width = 236 * 20
        #HEADER
        worksheet.row(4).height_mismatch = True
        worksheet.row(4).height = 360
        
        worksheet.write_merge(0, 1, 2, 5, self.env.company.name,styleheader)
        worksheet.write_merge(2, 2, 2, 5, 'Trial Balance',general)
        headerstring = 'From :' + str(self.date_from.strftime('%d %b %Y') or '') + ' To :' + str(self.date_to.strftime('%d %b %Y') or '')
        worksheet.write_merge(3, 3, 2, 5, headerstring,dateheader)
        
        dimension_res = ''
        month_res = ''
        if self.dimension_wise_project == 'dimension' :
            dimension_res = self.get_dimension_queries(option_list,lines)

        if self.dimension_wise_project == 'month' :
            month_res = self.get_monthwise_data(option_list,lines)
       
        #SUB-HEADER
        ColIndexes = {}
        row = 4
        calc = 10
        col = 9
        # colc = 4
        if self.show_dr_cr_separately:
            calc = 10
            col = 9
        else:
            calc = 8
            col = 7

        if self.dimension_wise_project == 'dimension':

            for analytic in AnalyticNames:
                dictval = {analytic:col}
                ColIndexes.update(dictval)
                dyna_col = worksheet.col(col)
                dyna_col.width = 236 * 20
                worksheet.write(row, col, analytic, analytic_st_col)
                # colc = col
                col+=1
                calc+=1

        elif self.dimension_wise_project == 'month':

            cur_date = self.date_from
            end = self.date_to
            while cur_date < end:
                cur_date_strf = str(cur_date.strftime('%b %y') or '')
                cur_date += relativedelta(months=1)
                dictval = {cur_date_strf:col}
                ColIndexes.update(dictval)
                dyna_col = worksheet.col(col)
                dyna_col.width = 236 * 20
                worksheet.write(row, col, cur_date_strf, analytic_st_col)
                # colc = col
                col+=1
                calc+=1

        worksheet.write(row, 0, 'Account Code', stylecolaccount)
        worksheet.write(row, 1, 'Account Name', stylecolaccount)
        opcol=2
        if self.show_dr_cr_separately:
            worksheet.write(row, opcol, 'Op.Debit', stylecolaccount)
            opcol+=1
            worksheet.write(row, opcol, 'Op.Credit', stylecolaccount)
            opcol+=1
        worksheet.write(row, opcol, 'Tr.Debit', stylecolaccount)
        opcol+=1
        worksheet.write(row, opcol, 'Tr.Credit', stylecolaccount)
        opcol+=1
        worksheet.write(row, opcol, 'Closing Debit', stylecolaccount)
        opcol+=1
        worksheet.write(row, opcol, 'Closing Credit', stylecolaccount)
        opcol+=1
        worksheet.write(row, opcol, 'Net Balance', stylecolaccount)
        opcol+=1
        row = 5

        dynamic_dimension_res = []
        Analyticvals = []
        final_list = []
        MonthVals = []
       
        for analytic in AnalyticNames:
            Analyticvals.append({analytic:0, 'class': 'number', 'no_format_name':0.0})

        fetch_monthwise_data = []
        cur_date = self.date_from
        end = self.date_to
        while cur_date < end:
            cur_date_strf = str(cur_date.strftime('%b %y') or '')
            cur_date += relativedelta(months=1)
            MonthVals.append({cur_date_strf:0, 'class': 'number', 'no_format_name':0.0})

        net_balance = 0.0
        for i in range(len(new_lines)):
            acc_balance = new_lines[i]['columns']
            listd = [list(c.values())[0] for c in acc_balance]
            first = listd[4]
            second = listd[5]
            net_balance = first - second
            acc_balance .append({'name':net_balance, 'class': 'number', 'no_format_name':net_balance})


            if self.dimension_wise_project == 'dimension':
                new_lines[i]['project'] = Analyticvals
            if self.dimension_wise_project == 'month':
                new_lines[i]['month'] = MonthVals
              
        if self.dimension_wise_project == 'dimension':
            for i in range(len(new_lines)):
                for dim in range(len(dimension_res)):
                    if dimension_res[dim]['account_code'] == new_lines[i]['code']:
                        new_lines[i]['project'] = dimension_res[dim]['columns']

        if self.dimension_wise_project == 'month':
            for i in range(len(new_lines)):
                for dim in range(len(month_res)):
                    if month_res[dim]['account_code'] == new_lines[i]['code']:
                        new_lines[i]['month'] = month_res[dim]['columns']
                       
        if not self.account_without_transaction :
            for i in range(len(new_lines)):
                v = 0
                acc_balance = new_lines[i]['columns']

                for j in acc_balance:
                    v +=list(j.values())[0]

                if v == 0:
                    continue
                else:
                    if self.show_dr_cr_separately:
                        acc_balance = acc_balance
                    else:
                        acc_balance = acc_balance[2:]
                    name = ''
                    code = ''
                    if not new_lines[i].get('code'):
                        name = new_lines[i]['name']
                        if isinstance(new_lines[i]['id'], int):              
                            worksheet.write(row, 0,'', mainheaderdata)
                            # col+=1
                            worksheet.write(row, 1 , name, mainheaderdata)
                            col = 2
                            for j in range(len(acc_balance)):
                                if acc_balance[j]['name'] == 0.0:
                                    worksheet.write(row, col, 00.0, mainheaderdata)
                                else:    
                                    worksheet.write(row, col, round((acc_balance[j]['name']),1), mainheaderdata)
                                col+=1
                            if self.dimension_wise_project == 'dimension':
                                projects = new_lines[i]['project']
                                if self.show_dr_cr_separately:
                                    col = 9
                                    for j in range(len(projects)):

                                        worksheet.write(row, col,round((list(projects[j].values())[0]),1), mainheaderdata)
                                        col+=1
                                    row+=1
                                else:
                                    col = 7
                                    for j in range(len(projects)):
                                        worksheet.write(row, col,round((list(projects[j].values())[0]),1), mainheaderdata)
                                        col+=1
                                    row+=1
                            elif self.dimension_wise_project == 'month':
                                months = new_lines[i]['month']
                                if self.show_dr_cr_separately:
                                    col = 9
                                    for j in range(len(months)):
                                        
                                        worksheet.write(row, col,round((list(months[j].values())[0]),1), mainheaderdata)
                                        col+=1
                                    row+=1
                                else:
                                    col = 7
                                    for j in range(len(months)):
                                        worksheet.write(row, col,round((list(months[j].values())[0]),1), mainheaderdata)
                                        col+=1
                                    row+=1

                            else:
                                row+=1
                        else:
                            worksheet.write(row, 0,'', mainheaders)
                            # col+=1
                            worksheet.write(row, 1 , name, mainheaders)
                            col = 2
                            for j in range(len(acc_balance)):
                                if acc_balance[j]['name'] == 0.0:
                                    worksheet.write(row, col, 00.0, mainheaders)
                                else:    
                                    worksheet.write(row, col, round((acc_balance[j]['name']),1), mainheaders)
                                col+=1
                            if self.dimension_wise_project == 'dimension':
                                projects = new_lines[i]['project']
                                if self.show_dr_cr_separately:
                                    col = 9
                                    for j in range(len(projects)):
                                        worksheet.write(row, col,round((list(projects[j].values())[0]),1), mainheaders)
                                        col+=1
                                    row+=1
                                else:
                                    col = 7
                                    for j in range(len(projects)):
                                        worksheet.write(row, col,round((list(projects[j].values())[0]),1), mainheaders)
                                        col+=1
                                    row+=1
                            elif self.dimension_wise_project == 'month':
                                months = new_lines[i]['month']
                                if self.show_dr_cr_separately:
                                    col = 9
                                    for j in range(len(months)):
                                        worksheet.write(row, col, round((list(months[j].values())[0]),1), mainheaders)
                                        col+=1
                                    row+=1
                                else:
                                    col = 7
                                    for j in range(len(months)):
                                        worksheet.write(row, col,round((list(months[j].values())[0]),1), mainheaders)
                                        col+=1
                                    row+=1

                            else:
                                row+=1

                    else:
                        name = new_lines[i]['name'].replace(new_lines[i]['code'], "")
                        code = new_lines[i]['code']
                        if isinstance(new_lines[i]['id'], int):
                            worksheet.write(row, 0,new_lines[i]['code'], mainheaderdata)
                            # col+=1
                            worksheet.write(row, 1, name, mainheaderdata)
                            col = 2
                            for j in range(len(acc_balance)):
                                if acc_balance[j]['name'] == 0.0:
                                    worksheet.write(row, col, 00.0, mainheaderdata)
                                else:    
                                    worksheet.write(row, col, round((acc_balance[j]['name']),1), mainheaderdata)
                                col+=1
                            if self.dimension_wise_project == 'dimension':
                                projects = new_lines[i]['project']
                                if self.show_dr_cr_separately:
                                    col = 9
                                    for j in range(len(projects)):
                                        worksheet.write(row, col,round((list(projects[j].values())[0]),1), mainheaderdata)
                                        col+=1
                                    row+=1
                                else:
                                    col = 7
                                    for j in range(len(projects)):
                                        worksheet.write(row, col,round((list(projects[j].values())[0]),1), mainheaderdata)
                                        col+=1
                                    row+=1
                            elif self.dimension_wise_project == 'month':
                                months = new_lines[i]['month']
                                if self.show_dr_cr_separately:
                                    col = 9
                                    for j in range(len(months)):
                                        worksheet.write(row, col, round((list(months[j].values())[0]),1), mainheaderdata)
                                        col+=1
                                    row+=1
                                else:
                                    col = 7
                                    for j in range(len(months)):
                                        worksheet.write(row, col, round((list(months[j].values())[0]),1), mainheaderdata)
                                        col+=1
                                    row+=1

                            else:
                                row+=1
                        else:
                            worksheet.write(row, 0,new_lines[i]['code'], mainheaders)
                            # col+=1
                            worksheet.write(row, 1, name, mainheaders)
                            col = 2
                            for j in range(len(acc_balance)):
                                if acc_balance[j]['name'] == 0.0:
                                    worksheet.write(row, col, 00.0, mainheaders)
                                else:    
                                    worksheet.write(row, col, round((acc_balance[j]['name']),1), mainheaders)
                                col+=1
                            if self.dimension_wise_project == 'dimension':
                                projects = new_lines[i]['project']
                                if self.show_dr_cr_separately:
                                    col = 9
                                    for j in range(len(projects)):
                                        worksheet.write(row, col,round((list(projects[j].values())[0]),1), mainheaders)
                                        col+=1
                                    row+=1
                                else:
                                    col = 7
                                    for j in range(len(projects)):
                                        worksheet.write(row, col,round((list(projects[j].values())[0]),1), mainheaders)
                                        col+=1
                                    row+=1
                            elif self.dimension_wise_project == 'month':
                                months = new_lines[i]['month']
                                if self.show_dr_cr_separately:
                                    col = 9
                                    for j in range(len(months)):
                                        worksheet.write(row, col,round((list(months[j].values())[0]),1), mainheaders)
                                        col+=1
                                    row+=1
                                else:
                                    col = 7
                                    for j in range(len(months)):
                                        worksheet.write(row, col,round((list(months[j].values())[0]),1), mainheaders)
                                        col+=1
                                    row+=1

                            else:
                                row+=1
                  
        else:
            for i in range(len(new_lines)):
                acc_balance = new_lines[i]['columns']
                if self.show_dr_cr_separately:
                    acc_balance = acc_balance
                else:
                    acc_balance = acc_balance[2:]
                
                name = ''
                code = ''
                if not new_lines[i].get('code'):
                    name = new_lines[i]['name']
                    if isinstance(new_lines[i]['id'], int):           
                        worksheet.write(row, 0,'', mainheaderdata)
                        # col+=1
                        worksheet.write(row, 1 , name, mainheaderdata)
                        col = 2
                        for j in range(len(acc_balance)):
                            if acc_balance[j]['name'] == 0.0:
                                worksheet.write(row, col, 00.0, mainheaderdata)
                            else:
                                worksheet.write(row, col, round((acc_balance[j]['name']),1), mainheaderdata)
                            col+=1
                        if self.dimension_wise_project == 'dimension':
                            projects = new_lines[i]['project']
                            if self.show_dr_cr_separately:
                                col = 9
                                for j in range(len(projects)):
                                    worksheet.write(row, col,round((list(projects[j].values())[0]),1), mainheaderdata)
                                    col+=1
                                row+=1
                            else:
                                col = 7
                                for j in range(len(projects)):
                                    worksheet.write(row, col, round((list(projects[j].values())[0]),1), mainheaderdata)
                                    col+=1
                                row+=1
                        elif self.dimension_wise_project == 'month':
                            months = new_lines[i]['month']
                            if self.show_dr_cr_separately:
                                col = 9
                                for j in range(len(months)):
                                    worksheet.write(row, col,round((list(months[j].values())[0]),1), mainheaderdata)
                                    col+=1
                                row+=1
                            else:
                                col = 7
                                for j in range(len(months)):
                                    worksheet.write(row, col,round((list(months[j].values())[0]),1), mainheaderdata)
                                    col+=1
                                row+=1

                        else:
                            row+=1
                    else:
                        worksheet.write(row, 0,'', mainheaders)
                        # col+=1
                        worksheet.write(row, 1 , name, mainheaders)
                        col = 2
                        for j in range(len(acc_balance)):
                            if acc_balance[j]['name'] == 0.0:
                                worksheet.write(row, col, 00.0, mainheaders)
                            else:    
                                worksheet.write(row, col, round((acc_balance[j]['name']),1), mainheaders)
                            col+=1
                        if self.dimension_wise_project == 'dimension':
                            projects = new_lines[i]['project']
                            if self.show_dr_cr_separately:
                                col = 9
                                for j in range(len(projects)):
                                    worksheet.write(row, col, round((list(projects[j].values())[0]),1), mainheaders)
                                    col+=1
                                row+=1
                            else:
                                col = 7
                                for j in range(len(projects)):
                                    worksheet.write(row, col,round((list(projects[j].values())[0]),1), mainheaders)
                                    col+=1
                                row+=1
                        elif self.dimension_wise_project == 'month':
                            months = new_lines[i]['month']
                            if self.show_dr_cr_separately:
                                col = 9
                                for j in range(len(months)):
                                    worksheet.write(row, col,round((list(months[j].values())[0]),1), mainheaders)
                                    col+=1
                                row+=1
                            else:
                                col = 7
                                for j in range(len(months)):
                                    worksheet.write(row, col, round((list(months[j].values())[0]),1), mainheaders)
                                    col+=1
                                row+=1

                        else:
                            row+=1


                else:
                    name = new_lines[i]['name'].replace(new_lines[i]['code'],"")
                    code = new_lines[i]['code']
                    if isinstance(new_lines[i]['id'], int):
                        worksheet.write(row, 0,new_lines[i]['code'], mainheaderdata) 
                        worksheet.write(row, 1, name, mainheaderdata)
                        col = 2
                        for j in range(len(acc_balance)):
                            if acc_balance[j]['name'] == 0.0:
                                worksheet.write(row, col, 00.0, mainheaderdata)
                            else:    
                                worksheet.write(row, col, round((acc_balance[j]['name']),1), mainheaderdata)
                            col+=1
                        if self.dimension_wise_project == 'dimension':
                            projects = new_lines[i]['project']
                            if self.show_dr_cr_separately:
                                col = 9
                                for j in range(len(projects)):
                                    worksheet.write(row, col, round((list(projects[j].values())[0]),1), mainheaderdata)
                                    col+=1
                                row+=1
                            else:
                                col = 7
                                for j in range(len(projects)):
                                    worksheet.write(row, col, round((list(projects[j].values())[0]),1), mainheaderdata)
                                    col+=1
                                row+=1
                        elif self.dimension_wise_project == 'month':
                            months = new_lines[i]['month']
                            if self.show_dr_cr_separately:
                                col = 9
                                for j in range(len(months)):
                                    worksheet.write(row, col,round((list(months[j].values())[0]),1), mainheaderdata)
                                    col+=1
                                row+=1
                            else:
                                col = 7
                                for j in range(len(months)):
                                    worksheet.write(row, col,round((list(months[j].values())[0]),1), mainheaderdata)
                                    col+=1
                                row+=1

                        else:
                            row+=1

                    else:
                        worksheet.write(row, 0,new_lines[i]['code'], mainheaders)
                        worksheet.write(row, 1, name, mainheaders)
                        col = 2
                        for j in range(len(acc_balance)):
                            if acc_balance[j]['name'] == 0.0 or acc_balance[j]['name'] == 0:
                                worksheet.write(row, col, 00.0, mainheaders)
                            else:    
                                worksheet.write(row, col, round((acc_balance[j]['name']),1), mainheaders)
                            col+=1
                            
                        if self.dimension_wise_project == 'dimension':
                            projects = new_lines[i]['project']
                            if self.show_dr_cr_separately:
                                col = 9
                                for j in range(len(projects)):
                                    worksheet.write(row, col,round((list(projects[j].values())[0]),1), mainheaders)
                                    col+=1
                                row+=1
                            else:
                                col = 7
                                for j in range(len(projects)):
                                    worksheet.write(row, col, round((list(projects[j].values())[0]),1), mainheaders)
                                    col+=1
                                row+=1
                        elif self.dimension_wise_project == 'month':
                            months = new_lines[i]['month']
                            if self.show_dr_cr_separately:
                                col = 9
                                for j in range(len(months)):
                                    worksheet.write(row, col,round((list(months[j].values())[0]),1), mainheaders)
                                    col+=1
                                row+=1
                            else:
                                col = 7
                                for j in range(len(months)):
                                    worksheet.write(row, col,round((list(months[j].values())[0]),1), mainheaders)
                                    col+=1
                                row+=1

                        else:
                            row+=1


        row+=2
        buffer = io.BytesIO()
        workbook.save(buffer)
        export_id = self.env['trial.balance.excel'].create(
                        {'excel_file': base64.encodestring(buffer.getvalue()), 'file_name': filename})
        buffer.close()
    
        return {
            'name': form_name,
            'view_mode': 'form',
            'res_id': export_id.id,
            'res_model': 'trial.balance.excel',
            'view_mode': 'form',
            'type': 'ir.actions.act_window',
            'target': 'new',
        }

    @api.model
    def get_options_sum_balance(self, options):
        new_options = options.copy()
        fiscalyear_dates = self.env.company.compute_fiscalyear_dates(fields.Date.from_string(new_options['date']['date_from']))
        new_options['date'] = {
            'mode': 'range',
            'date_from': fiscalyear_dates['date_from'].strftime('%Y-%m-%d'),
            'date_to': options['date']['date_to'],
        }
        return new_options    

    @api.model
    def get_options_unaffected_earnings(self, options):
        new_options = options.copy()
        fiscalyear_dates = self.env.company.compute_fiscalyear_dates(fields.Date.from_string(options['date']['date_from']))
        new_date_to = fiscalyear_dates['date_from'] - timedelta(days=1)
        new_options['date'] = {
            'mode': 'single',
            'date_to': new_date_to.strftime('%Y-%m-%d'),
        }
        return new_options

    @api.model
    def get_options_initial_balance(self, options):
        new_options = options.copy()
        fiscalyear_dates = self.env.company.compute_fiscalyear_dates(fields.Date.from_string(options['date']['date_from']))
        new_date_to = fields.Date.from_string(new_options['date']['date_from']) - timedelta(days=1)
        new_options['date'] = {
            'mode': 'range',
            'date_from': fiscalyear_dates['date_from'].strftime('%Y-%m-%d'),
            'date_to': new_date_to.strftime('%Y-%m-%d'),
        }
        return new_options    
    @api.model
    def get_options_date_domain(self, options):
        def create_date_domain(options_date):
            date_field = options_date.get('date_field', 'date')
            domain = [('company_id', 'in', [self.env.company.id]),(date_field, '<=', options_date['date_to'])]
            if options_date['mode'] == 'range':
                strict_range = options_date.get('strict_range')
                if not strict_range:
                    domain += [
                        '|',
                        (date_field, '>=', options_date['date_from']),
                        ('account_id.user_type_id.include_initial_balance', '=', True),
                        ('move_id.state', '=', 'posted'),
                    ]
                else:
                    domain += [(date_field, '>=', options_date['date_from'])]
            else:
                domain +=[('move_id.state', '=', 'posted')]
            return domain

        if not options.get('date'):
            return []
        return create_date_domain(options['date']) 

    @api.model
    def query_get(self, options, domain=None):
        domain = self.get_options_date_domain(options) + (domain or [])
        self.env['account.move.line'].check_access_rights('read')

        query = self.env['account.move.line']._where_calc(domain)
        # Wrap the query with 'company_id IN (...)' to avoid bypassing company access rights.
        self.env['account.move.line']._apply_ir_rules(query)

        return query.get_sql()               

    def get_all_queries(self,option_list,expanded_account=None):
        queries = []
        params = []
        queries = []
        dynamic_queries = []
        option_list = option_list.copy()
        user_company = self.env.company
        user_currency = user_company.currency_id
        companies = user_company
        currency_rates = {user_currency.id: 1.0}
        conversion_rates = []
        for company in companies:
            conversion_rates.append((
                company.id,
                currency_rates[user_company.currency_id.id] / currency_rates[company.currency_id.id],
                user_currency.decimal_places,
            ))

        currency_table = ','.join('(%s, %s, %s)' % args for args in conversion_rates)
        ct_query = '(VALUES %s) AS currency_table(company_id, rate, precision)' % currency_table

        # ============================================
        # 1) Get sums for all accounts.
        # ============================================

        domain = [('account_id', '=', expanded_account.id)] if expanded_account else []

        for i, options_period in enumerate(option_list):
            new_options = self.get_options_sum_balance(options_period)
            tables, where_clause, where_params = self.query_get(new_options, domain=domain)
            params += where_params
            queries.append('''
                SELECT
                    account_move_line.account_id                            AS groupby,
                    'sum'                                                   AS key,
                    MAX(account_move_line.date)                            AS max_date,
                    %s                                                      AS period_number,
                    COALESCE(SUM(account_move_line.amount_currency), 0.0)   AS amount_currency,
                    SUM(ROUND(account_move_line.debit * currency_table.rate, currency_table.precision))   AS debit,
                    SUM(ROUND(account_move_line.credit * currency_table.rate, currency_table.precision))  AS credit,
                    SUM(ROUND(account_move_line.balance * currency_table.rate, currency_table.precision)) AS balance
                FROM %s
                LEFT JOIN %s ON currency_table.company_id = account_move_line.company_id
                WHERE %s
                GROUP BY account_move_line.account_id
            ''' % (i, tables, ct_query, where_clause))
           
        # ============================================
        # 2) Get sums for the unaffected earnings.
        # ============================================

        domain = [('account_id.user_type_id.include_initial_balance', '=', False)]
        # if expanded_account:
        #     domain.append(('company_id', '=', expanded_account.company_id.id))

        # Compute only the unaffected earnings for the oldest period.
        i = len(option_list) - 1
        options_period = option_list[-1]
        new_options = self.get_options_unaffected_earnings(options_period)
        tables, where_clause, where_params = self.query_get(new_options, domain=domain)
        params += where_params
        queries.append('''
            SELECT
                account_move_line.company_id                            AS groupby,
                'unaffected_earnings'                                   AS key,
                NULL                                                    AS max_date,
                %s                                                      AS period_number,
                COALESCE(SUM(account_move_line.amount_currency), 0.0)   AS amount_currency,
                SUM(ROUND(account_move_line.debit * currency_table.rate, currency_table.precision))   AS debit,
                SUM(ROUND(account_move_line.credit * currency_table.rate, currency_table.precision))  AS credit,
                SUM(ROUND(account_move_line.balance * currency_table.rate, currency_table.precision)) AS balance
            FROM %s
            LEFT JOIN %s ON currency_table.company_id = account_move_line.company_id
            WHERE %s
            GROUP BY account_move_line.company_id
        ''' % (i, tables, ct_query, where_clause))
        # ============================================
        # 3) Get sums for the initial balance.
        # ============================================
        domain = None
        if domain is None:
            domain = []
        if domain is not None:
            for i, options_period in enumerate(option_list):
                new_options = self.get_options_initial_balance(options_period)
                tables, where_clause, where_params = self.query_get(new_options, domain=domain)
                params += where_params
                queries.append('''
                    SELECT
                        account_move_line.account_id                            AS groupby,
                        'initial_balance'                                       AS key,
                        NULL                                                    AS max_date,
                        %s                                                      AS period_number,
                        COALESCE(SUM(account_move_line.amount_currency), 0.0)   AS amount_currency,
                        SUM(ROUND(account_move_line.debit * currency_table.rate, currency_table.precision))   AS debit,
                        SUM(ROUND(account_move_line.credit * currency_table.rate, currency_table.precision))  AS credit,
                        SUM(ROUND(account_move_line.balance * currency_table.rate, currency_table.precision)) AS balance
                    FROM %s
                    LEFT JOIN %s ON currency_table.company_id = account_move_line.company_id
                    WHERE %s
                    GROUP BY account_move_line.account_id
                ''' % (i, tables, ct_query, where_clause))

        return ' UNION ALL '.join(queries), params

    def get_dimesnsion_hierarchy_lines(self,values, depth=1):
        
        lines = []
        sum_sum_columns = []
        sum_name_coumns = []
        # unfold_all = self.env.context.get('print_mode') and len(options.get('unfolded_lines')) == 0
        for base_line in values.get('lines', []):
            lines.append(base_line)

            sum_sum_columns.append([c.get('no_format_name',list(c.values())[0]) for c in base_line['columns']])
        
        # For the last iteration, there might not be the children key (see add_line_to_hierarchy)
        for key in sorted(values.get('children', {}).keys()):
            
            sum_columns, sub_lines = self.get_dimesnsion_hierarchy_lines(values['children'][key], depth=values['depth'])
            sum_columns_bal = [c if not isinstance(c, str) else c for c in sum_columns]
            AnalyticAccount = self.analytic_account_ids
            ac_names = [i.name for i in AnalyticAccount]

            dictionary = dict(zip(ac_names, sum_columns_bal))
            dicts = []

            for i,s in dictionary.items():
                dicts.append({i:s,'class': 'number', 'no_format_name':s})
            id = 'hierarchy_' + key[1]
            acc_code = key[1].split(" ")[0]
            header_line = {
                'id': id,
                'account_code':acc_code,
                'name': key[1] if len(key[1]) < 60 else key[1][:60] + '...',  # second member of the tuple
                'title_hover': key[1],
                'unfoldable': True,
                'level': values['depth'],
                'parent_id': values['parent_id'],
                'columns': dicts,
            }
            if key[0] == self.LEAST_SORT_PRIO:
                header_line['style'] = 'font-style:italic;'
            lines += [header_line] + sub_lines
            sum_sum_columns.append(sum_columns)
        return self.merge_columns(sum_sum_columns), lines


    def dimension_group_line_calculation(self,lines):

        accounts_hierarchy = {}
        no_group_lines = []
        new_lines = []
        for line in lines + [None]:
            is_grouped_by_account = line and (line.get('caret_options') == 'account.account' or line.get('account_id'))
            if not is_grouped_by_account or not line:
                no_group_hierarchy = {}
                for no_group_line in no_group_lines:
                    codes = [('root', str(line.get('parent_id')) or 'root') if line else 'root', (self.LEAST_SORT_PRIO, _('(No Group)'))]
                    if not accounts_hierarchy:
                        account = self.get_account(no_group_line.get('account_id', no_group_line.get('id')))
                        codes = [('root', line and str(line.get('parent_id')) or 'root')] + self.get_account_codes(account)
                    self.add_line_to_hierarchy(no_group_line, codes, no_group_hierarchy, line and line.get('level') or 0 + 1)
                no_group_lines = []
                
                self.deep_merge_dict(no_group_hierarchy, accounts_hierarchy)
                # Merge the newly created hierarchy with existing lines.
                if accounts_hierarchy:
                   
                    new_lines += self.get_dimesnsion_hierarchy_lines(accounts_hierarchy)[1]
                    # new_dimension_list.append(new_lines)
                    accounts_hierarchy = {}

                if line:
                    new_lines.append(line)
                continue

            # Exclude lines having no group.
            account = self.get_account(line.get('account_id', line.get('id')))

            if not account:
                no_group_lines.append(line)
                continue
            codes = [('root', str(line.get('parent_id')) or 'root')] + self.get_account_codes(account)
          
            self.add_line_to_hierarchy(line, codes, accounts_hierarchy,line.get('level', 0) + 1)
        return new_lines
        
    def get_month_hierarchy_lines(self,values, depth=1):
        lines = []
        sum_sum_columns = []
        sum_name_coumns = []
        # unfold_all = self.env.context.get('print_mode') and len(options.get('unfolded_lines')) == 0
        for base_line in values.get('lines', []):
            lines.append(base_line)

            sum_sum_columns.append([c.get('no_format_name',list(c.values())[0]) for c in base_line['columns']])
        
        # For the last iteration, there might not be the children key (see add_line_to_hierarchy)
        for key in sorted(values.get('children', {}).keys()):
            
            sum_columns, sub_lines = self.get_month_hierarchy_lines(values['children'][key], depth=values['depth'])
            sum_columns_bal = [c if not isinstance(c, str) else c for c in sum_columns]
            
            fetch_monthwise_data = []
            cur_date = self.date_from
            end = self.date_to
            while cur_date < end:
                cur_date_strf = str(cur_date.strftime('%b %y') or '')
                cur_date += relativedelta(months=1)
                fetch_monthwise_data.append(cur_date_strf)

            dictionary = dict(zip(fetch_monthwise_data, sum_columns_bal))
            monthdicts = []
            for i,s in dictionary.items():
                monthdicts.append({i:s,'class': 'number', 'no_format_name':s})
            id = 'hierarchy_' + key[1]
            acc_code = key[1].split(" ")[0]
            header_line = {
                'id': id,
                'account_code':acc_code,
                'name': key[1] if len(key[1]) < 60 else key[1][:60] + '...',  # second member of the tuple
                'title_hover': key[1],
                'unfoldable': True,
                'level': values['depth'],
                'parent_id': values['parent_id'],
                'columns': monthdicts,
            }
            if key[0] == self.LEAST_SORT_PRIO:
                header_line['style'] = 'font-style:italic;'

            lines += [header_line] + sub_lines

            sum_sum_columns.append(sum_columns)

        return self.merge_columns(sum_sum_columns), lines

    def month_group_line_calculation(self,lines):

        accounts_hierarchy = {}
        no_group_lines = []
        new_lines = []

        for line in lines + [None]:

            is_grouped_by_account = line and (line.get('caret_options') == 'account.account' or line.get('account_id'))
            if not is_grouped_by_account or not line:
                no_group_hierarchy = {}
                for no_group_line in no_group_lines:
                    codes = [('root', str(line.get('parent_id')) or 'root') if line else 'root', (self.LEAST_SORT_PRIO, _('(No Group)'))]
                    if not accounts_hierarchy:

                        account = self.get_account(no_group_line.get('account_id', no_group_line.get('id')))
                        codes = [('root', line and str(line.get('parent_id')) or 'root')] + self.get_account_codes(account)
                    self.add_line_to_hierarchy(no_group_line, codes, no_group_hierarchy, line and line.get('level') or 0 + 1)
                no_group_lines = []

                self.deep_merge_dict(no_group_hierarchy, accounts_hierarchy)
                # Merge the newly created hierarchy with existing lines.
                if accounts_hierarchy:

                    new_lines += self.get_month_hierarchy_lines(accounts_hierarchy)[1]

                    # new_dimension_list.append(new_lines)
                    accounts_hierarchy = {}

                if line:
                    new_lines.append(line)

                continue

            # Exclude lines having no group.
            account = self.get_account(line.get('account_id', line.get('id')))

            if not account.group_id.id:
                no_group_lines.append(line)
                continue
            codes = [('root', str(line.get('parent_id')) or 'root')] + self.get_account_codes(account)
          
            self.add_line_to_hierarchy(line, codes, accounts_hierarchy,line.get('level', 0) + 1)
        return new_lines
          
    def get_dimension_queries(self,option_list,lines):

        options_period = option_list[-1]
        AnalyticAccountIds = options_period['analytic_accounts']
        dateFrom = self.date_from
        dateTo = self.date_to
        Status = ['posted']
        MoveLineIds = []
        AccountId = []
        account_list = []
        main_list = []
        first_list =[]
        mainDict = []
        second_list = []
        column1 = []
        new_list = []
        for i in range(len(lines)):
            if isinstance(lines[i]['id'], int):
                AccountId.append(lines[i]['id'])
            self.env.cr.execute("""
                SELECT aml.date as date,
                       aml.debit as debit,
                       aml.credit as credit,
                       aa.name as analytic,
                       aml.account_id as account_id,
                       a.name as account_name,
                       a.group_id as group_id,
                       ag.name as group_name,
                       ag.code_prefix as group_code,
                       a.code as account_code,
                       aml.id as movelineid
                FROM account_move_line aml
                LEFT JOIN account_move am ON (am.id=aml.move_id)
                LEFT JOIN account_account a ON (aml.account_id=a.id)
                LEFT JOIN account_group ag ON (a.group_id=ag.id)
                LEFT JOIN account_analytic_account aa ON (aa.id=aml.analytic_account_id)
                WHERE (aml.date >= %s) AND
                    (aml.date <= %s) AND
                    (aml.account_id in %s) AND
                    (aml.analytic_account_id in %s) AND
                    (am.state in %s) ORDER BY aml.account_id""",
                (str(dateFrom) + ' 00:00:00', str(dateTo) + ' 23:59:59', tuple(AccountId),tuple(AnalyticAccountIds), tuple(Status),))
            MoveLineIds = self.env.cr.fetchall()
        
        if MoveLineIds:
                for ml in MoveLineIds:
                    date = ml[0]
                    acount_debit = ml[1]
                    account_credit = ml[2]
                    analytic_account_id = ml[3]
                    account_id = ml[4]
                    account_name = ml[5]
                    group_id = ml[6]
                    group_name = ml[7]
                    group_code = ml[8]
                    account_code = ml[9]
                    Balance = 0.0
                    Balance = Balance + (acount_debit - account_credit)
                    Vals = {
                            'account_id':account_id,
                            'account_name':account_name,
                            'group_id':group_id,
                            'group_name':group_name,
                            'group_code':group_code,                      
                            'balance': Balance or 0.0,
                            'account_debit':acount_debit,
                            'account_credit':account_credit,
                            'analytic_account_id':analytic_account_id,
                            'date':date,
                            'account_code':account_code,
                            }   
                    mainDict.append(Vals)
        
        for i in range(0,len(mainDict)):
            if (mainDict[i]['account_name'],mainDict[i]['analytic_account_id']) not in account_list:
                main_list.append({
                                  'id':mainDict[i]['account_id'],
                                  'account_name':mainDict[i]['account_name'],
                                  'group_id':mainDict[i]['group_id'],
                                  'group_name':mainDict[i]['group_name'],
                                  'group_code': mainDict[i]['group_code'],
                                  'analytic_account_id':mainDict[i]['analytic_account_id'],
                                  'debit': mainDict[i]['account_debit'],
                                  'credit': mainDict[i]['account_credit'],
                                  'account_code': mainDict[i]['account_code'],
                                  'balance': mainDict[i]['account_debit'] - mainDict[i]['account_credit'] or 00.00,
                                  })
                account_list.append((mainDict[i]['account_name'],mainDict[i]['analytic_account_id']))
            else:
                first_list.append({
                                  'id':mainDict[i]['account_id'],
                                  'account_name':mainDict[i]['account_name'],
                                  'group_id':mainDict[i]['group_id'],
                                  'group_name':mainDict[i]['group_name'],
                                  'group_code': mainDict[i]['group_code'],
                                  'analytic_account_id':mainDict[i]['analytic_account_id'],
                                  'debit': mainDict[i]['account_debit'],
                                  'credit': mainDict[i]['account_credit'],
                                  'account_code': mainDict[i]['account_code'],
                                  'balance': mainDict[i]['account_debit'] - mainDict[i]['account_credit'] or 00.00,
                                  })
        for j in range(0,len(main_list)):     
            for k in range(0,len(first_list)):
                if first_list[k]['account_name'] == main_list[j]['account_name'] and first_list[k]['analytic_account_id'] == main_list[j]['analytic_account_id']:
                    main_list[j]['debit'] =  main_list[j]['debit'] + first_list[k]['debit']
                    main_list[j]['credit'] = main_list[j]['credit'] + first_list[k]['credit']
                    main_list[j]['balance'] = main_list[j]['debit'] - main_list[j]['credit']
            
        for i in range(0,len(main_list)):
            if main_list[i]['id'] not in second_list:
                new_list.append(main_list[i])
                second_list.append(main_list[i]['id'])
        
        ana_id = self.env['account.analytic.account'].browse(AnalyticAccountIds)
        ac_names = [i.name for i in ana_id]
        a3 = ''
        res2 =''
        totaldebit = 0.0
        totalcredit = 0.0
        totalbalance = 0.0
        total_balance = 0.0
        analytic_list = []
        new_analytic_list = []
        total_balance_list = []
        total_list = []
        another_analytic_list = []
        third_income_lists = []
        for j in range(0,len(new_list)):
            for k in range(0,len(main_list)):
                if new_list[j]['id'] == main_list[k]['id']:
                    column1.append({main_list[k]['analytic_account_id']:main_list[k]['balance'],  'class': 'number', 'no_format_name': main_list[k]['balance']})
                    a1 = [(list(c.keys())[0]) for c in column1]
                    res = column1 + [{i:0, 'class': 'number', 'no_format_name': 0.0} for i in ac_names if i not in a1]
                    res2 = sorted(res, key=lambda d: sorted(d.items()))
                    new_list[j]['columns'] = res2
                    new_list[j]['caret_options'] = 'account.account'
                    new_list[j]['level'] = 0
                    new_list[j]['parent_id'] = 'hierarchy_' + str(main_list[k]['group_code']) + str(" ") + str(main_list[k]['group_name'])
                else:
                   column1.clear()

        for s in range(0,len(new_list)):
            totalcolumn = new_list[s]['columns']
            listd = [list(c.values())[0] for c in totalcolumn]
            third_income_lists.append(listd)

        finalincomelist = [sum(i) for i in zip(*third_income_lists)]
        total_balance_list = dict(zip(ac_names, finalincomelist))
        finalbalancedict = [{k:v, 'class': 'number', 'no_format_name': v} for k,v in total_balance_list.items()]
          
        new_list.append({
             'id': 'grouped_accounts_total',
             'account_code': 'group_code',
             'name': _('Total'),
             'class': 'total',
             'columns': finalbalancedict,
             'level': 1,
        })

        return self.dimension_group_line_calculation(new_list)

    def get_monthwise_data(self,option_list,lines):
        options_period = option_list[-1]
        AnalyticAccountIds = options_period['analytic_accounts']
        dateFrom = self.date_from
        dateTo = self.date_to
        Status = ['posted']
        MoveLineIds = []
        AccountId = []
        account_list = []
        main_list = []
        first_list =[]
        mainDict = []
        second_list = []
        column1 = []
        new_list = []
        for i in range(len(lines)):
            if isinstance(lines[i]['id'], int):
                AccountId.append(lines[i]['id'])
            self.env.cr.execute("""
                SELECT aml.date as date,
                       aml.debit as debit,
                       aml.credit as credit,
                       aml.account_id as account_id,
                       a.name as account_name,
                       a.group_id as group_id,
                       ag.name as group_name,
                       ag.code_prefix as group_code,
                       a.code as account_code,
                       aml.id as movelineid
                FROM account_move_line aml
                LEFT JOIN account_move am ON (am.id=aml.move_id)
                LEFT JOIN account_account a ON (aml.account_id=a.id)
                LEFT JOIN account_group ag ON (a.group_id=ag.id)
                WHERE (aml.date >= %s) AND
                    (aml.date <= %s) AND
                    (aml.account_id in %s) AND
                    (am.state in %s) ORDER BY aml.account_id""",
                (str(dateFrom) + ' 00:00:00', str(dateTo) + ' 23:59:59', tuple(AccountId),tuple(Status),))
            MoveLineIds = self.env.cr.fetchall()
        
        if MoveLineIds:
            for ml in MoveLineIds:
                date = ml[0]
                acount_debit = ml[1]
                account_credit = ml[2]
                account_id = ml[3]
                account_name = ml[4]
                group_id = ml[5]
                group_name = ml[6]
                group_code = ml[7]
                account_code = ml[8]
                Balance = 0.0
                Balance = Balance + (acount_debit - account_credit)
                Vals = {
                        'account_id':account_id,
                        'account_name':account_name,
                        'group_id':group_id,
                        'group_name':group_name,
                        'group_code':group_code,                      
                        'balance': Balance or 0.0,
                        'account_debit':acount_debit,
                        'account_credit':account_credit,
                        'date':date,
                        'account_code':account_code,
                        }   
                mainDict.append(Vals)

        for i in range(0,len(mainDict)):
            if (mainDict[i]['account_name'],mainDict[i]['date'].strftime("%b %y")) not in account_list:
                main_list.append({
                                  'id':mainDict[i]['account_id'],
                                  'account_name':mainDict[i]['account_name'],
                                  'account_code': mainDict[i]['account_code'],
                                  'group_id':mainDict[i]['group_id'],
                                  'group_name':mainDict[i]['group_name'],
                                  'group_code': mainDict[i]['group_code'],
                                  'debit': mainDict[i]['account_debit'],
                                  'credit': mainDict[i]['account_credit'],
                                  'balance': mainDict[i]['account_debit'] - mainDict[i]['account_credit'] or 00.00,
                                  'month': mainDict[i]['date'].strftime("%b %y")                   
                                  })

                account_list.append((mainDict[i]['account_name'],mainDict[i]['date'].strftime("%b %y")))     
            else:
                first_list.append({
                                  'id':mainDict[i]['account_id'],
                                  'account_name':mainDict[i]['account_name'],
                                  'account_code': mainDict[i]['account_code'],
                                  'group_id':mainDict[i]['group_id'],
                                  'group_name':mainDict[i]['group_name'],
                                  'group_code': mainDict[i]['group_code'],
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


        for i in range(0,len(main_list)):
            if main_list[i]['id'] not in second_list:
                new_list.append(main_list[i])
                second_list.append(main_list[i]['id'])
        
        fetch_monthwise_data = []
        cur_date = dateFrom

        end = dateTo
        while cur_date < end:
            cur_date_strf = str(cur_date.strftime('%b %y') or '')
            cur_date += relativedelta(months=1)
            fetch_monthwise_data.append(cur_date_strf)

        a3 = ''
        res2 =''
        totaldebit = 0.0
        totalcredit = 0.0
        totalbalance = 0.0
        total_balance = 0.0
        month_list = []
        new_month_list = []
        total_balance_list = []
        total_list = []
        another_month_list = []
        third_income_lists = []
        for j in range(0,len(new_list)):
            for k in range(0,len(main_list)):
                if new_list[j]['id'] == main_list[k]['id']:
                    column1.append({main_list[k]['month']:main_list[k]['balance'], 'class': 'number', 'no_format_name': main_list[k]['balance']})
                    a1 = [(list(c.keys())[0]) for c in column1]
                    res = column1 + [{i:0.0, 'class': 'number', 'no_format_name': 0.0} for i in fetch_monthwise_data if i not in a1]
                    res2 = sorted(res, key = lambda ele: fetch_monthwise_data.index(list(ele.keys())[0]))
                    new_list[j]['columns'] = res2
                    new_list[j]['caret_options'] = 'account.account'
                    new_list[j]['level'] = 0
                    new_list[j]['parent_id'] = 'hierarchy_' + str(main_list[k]['group_code']) + str(" ") + str(main_list[k]['group_name'])
                else:
                   column1.clear()


        for s in range(0,len(new_list)):
            totalcolumn = new_list[s]['columns']
            listd = [list(c.values())[0] for c in totalcolumn]
            third_income_lists.append(listd)

        finalincomelist = [sum(i) for i in zip(*third_income_lists)]
        total_balance_list = dict(zip(fetch_monthwise_data, finalincomelist))
        finalbalancedict = [{k:v, 'class': 'number', 'no_format_name': v} for k,v in total_balance_list.items()]
        

        new_list.append({
             'id': 'grouped_accounts_total',
             'account_code': 'group_code',
             'name': _('Total'),
             'class': 'total',
             'columns': finalbalancedict,
             'level': 1,
        })
        return self.month_group_line_calculation(new_list)
      
class trial_balance_export_excel(models.TransientModel):
    _name= "trial.balance.excel"
    _description = "Trial Balance Excel Report"

    excel_file = fields.Binary('Report for Trial Balance')
    file_name = fields.Char('File', size=64)
