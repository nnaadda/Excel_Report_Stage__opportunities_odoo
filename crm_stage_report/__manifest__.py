{
    'name': 'CRM Stage Report',
    'version': '1.0.0',
    'category': 'Sales/CRM',
    'summary': 'Generate time-based stage reports for CRM leads',
    'description': '''
        CRM Stage Report Module
        =======================
        
        This module provides simplified stage reporting with:
        - Monthly stage summaries
        - Lead count per stage
        - Sales person and team tracking
        - Excel report generation
        - Year-based filtering
    ''',
    'author': 'Your Company',
    'depends': ['base', 'crm', 'mail'],
    'data': [
        'security/ir.model.access.csv',
        'views.xml',
    ],
    'installable': True,
    'auto_install': False,
    'application': False,
    'license': 'LGPL-3',
}