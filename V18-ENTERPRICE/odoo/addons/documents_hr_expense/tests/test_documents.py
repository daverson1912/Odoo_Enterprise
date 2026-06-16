# -*- coding: utf-8 -*-

import base64
from odoo import Command
from odoo.exceptions import UserError
from odoo.tests.common import TransactionCase
from odoo.tools import mute_logger

TEXT = base64.b64encode(bytes("workflow bridge project", 'utf-8'))


class TestCaseDocumentsBridgeExpense(TransactionCase):

    @classmethod
    @mute_logger('odoo.addons.documents.models.documents_document')
    def setUpClass(cls):
        super().setUpClass()

        folder_internal = cls.env.ref('documents.document_internal_folder')
        folder_internal.action_update_access_rights(access_internal='edit')

        cls.documents_user = cls.env['res.users'].create({
            'name': "aaadocuments test basic user",
            'login': "aadtbu",
            'email': "aadtbu@yourcompany.com",
            'groups_id': [Command.set([cls.env.ref('documents.group_documents_user').id])],
        })
        cls.attachment_txt = cls.env['documents.document'].with_user(cls.documents_user).create({
            'datas': 'JVBERi0gRmFrZSBQREYgY29udGVudA==',
            'name': 'file.pdf',
            'mimetype': 'application/pdf',
            'folder_id': folder_internal.id,
        })

    def test_create_document_to_expense(self):
        """
        Makes sure the hr expense is created from the document.

        Steps:
            - Create user with employee
            - Create attachment
            - Performed action 'Create a Expense'
            - Check if the expense is created
            - Check the res_model of the document

        """
        self.documents_user.action_create_employee()  # Employee is mandatory in expense

        self.assertEqual(self.attachment_txt.res_model, 'documents.document', "The default res model of an attachment is documents.document.")
        self.attachment_txt.with_user(self.documents_user).document_hr_expense_create_hr_expense()
        self.assertEqual(self.attachment_txt.res_model, 'hr.expense', "The attachment model is updated.")

        expense = self.env['hr.expense'].search([('id', '=', self.attachment_txt.res_id)])
        self.assertTrue(expense.exists(), 'expense sholud be created.')
        self.assertEqual(self.attachment_txt.res_id, expense.id, "Expense should be linked to attachment")

    def test_create_document_to_expense_without_employee(self):
        """
        Make sure UserError is raised when creating expense from document
        while the current user is not linked to an employee.
        """
        with self.assertRaisesRegex(UserError, "You must be linked to an employee to create an expense."):
            self.attachment_txt.with_user(self.documents_user).document_hr_expense_create_hr_expense()

        expense = self.env['hr.expense'].search([('id', '=', self.attachment_txt.res_id)])
        self.assertFalse(expense.exists())
