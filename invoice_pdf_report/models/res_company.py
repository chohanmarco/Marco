# -*- coding: utf-8 -*-
##############################################################################
#                                                                            #
# Part of Caret IT Solutions Pvt. Ltd. (Website: www.caretit.com).           #
# See LICENSE file for full copyright and licensing details.                 #
#                                                                            #
##############################################################################
from odoo import models, fields, api, _


class ResCompany(models.Model):
    _inherit = "res.company"


    stamp = fields.Binary(string="Stamp")
