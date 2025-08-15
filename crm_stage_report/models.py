from odoo import models, fields, api
import io
import xlsxwriter
from datetime import datetime, date, timedelta
import base64
from dateutil.relativedelta import relativedelta


class CrmLead(models.Model):
    _inherit = 'crm.lead'

    def _get_stage_visit_counts(self, date_from=None, date_to=None):
        """
        Return a dict mapping stage_id -> set(lead_id) for leads that visited that stage
        within the provided date range. Each lead is counted only once per stage.
        """
        Tracking = self.env['mail.tracking.value']

        # Find field_id for crm.lead.stage_id
        stage_field = self.env['ir.model.fields'].search([
            ('model', '=', 'crm.lead'),
            ('name', '=', 'stage_id')
        ], limit=1)
        if not stage_field:
            return {}

        domain = [
            ('field_id', '=', stage_field.id),
            ('mail_message_id.model', '=', 'crm.lead'),
        ]
        if date_from:
            domain.append(('create_date', '>=', date_from))
        if date_to:
            domain.append(('create_date', '<=', date_to))

        trackings = Tracking.search(domain)

        # Use dict of sets to ensure uniqueness
        stage_visits = {}
        for tracking in trackings:
            stage_id = tracking.new_value_integer
            if not stage_id and tracking.new_value_char:
                stage = self.env['crm.stage'].search(
                    [('name', '=', tracking.new_value_char)],
                    limit=1
                )
                stage_id = stage.id if stage else False

            lead_id = tracking.mail_message_id.res_id if tracking.mail_message_id else False

            if stage_id and lead_id:
                # Store in a set for uniqueness
                if stage_id not in stage_visits:
                    stage_visits[stage_id] = set()
                stage_visits[stage_id].add(lead_id)

        return stage_visits


    @api.model
    def action_generate_stage_summary_report(self, year=None, month=None):
        """Generate stage report (Excel) where lead_count = number of unique leads that visited the stage
           during the selected month (or per month when month is None = All Months)."""
        if not year:
            year = datetime.now().year

        # Prepare Excel file in memory
        output = io.BytesIO()
        workbook = xlsxwriter.Workbook(output, {'in_memory': True})

        # Define formats
        header_format = workbook.add_format({
            'bold': True,
            'bg_color': '#4472C4',
            'font_color': 'white',
            'border': 1,
            'text_wrap': True,
            'align': 'center'
        })
        data_format = workbook.add_format({'border': 1, 'text_wrap': True, 'valign': 'top'})
        number_format = workbook.add_format({'border': 1, 'align': 'center'})

        # Sheet name (truncate to 31 chars to be safe)
        if month:
            sheet_name = f"Stage Summary {date(year, month, 1).strftime('%B %Y')}"
        else:
            sheet_name = f"Stage Summary {year}"
        sheet = workbook.add_worksheet(sheet_name[:31])

        # Headers
        headers = [
            "Year",
            "Month",
            "Stage Name",
            "Lead Count",
            "Sales Persons (Active)",
            "Sales Teams (Active)"
        ]
        for col, header in enumerate(headers):
            sheet.write(0, col, header, header_format)

        # Get data for one month or for all months
        if month:
            stage_data = self._get_stage_data_by_tracking(year, month)
        else:
            stage_data = []
            for m in range(1, 13):
                monthly_data = self._get_stage_data_by_tracking(year, m)
                # ensure year/month/month_name fields exist for Excel write
                for rec in monthly_data:
                    rec['year'] = year
                    rec['month'] = m
                    rec['month_name'] = date(year, m, 1).strftime('%B')
                stage_data.extend(monthly_data)

        # Write data rows
        row = 1
        for data in stage_data:
            sheet.write(row, 0, data.get('year', year), number_format)
            sheet.write(row, 1, data.get('month_name', 'All'), data_format)
            sheet.write(row, 2, data['stage_name'], data_format)
            sheet.write(row, 3, data['lead_count'], number_format)
            sheet.write(row, 4, data['sales_persons'], data_format)
            sheet.write(row, 5, data['sales_teams'], data_format)
            row += 1

        # Column widths & filter
        column_widths = [8, 12, 30, 12, 40, 40]
        for col, width in enumerate(column_widths):
            sheet.set_column(col, col, width)
        sheet.autofilter(0, 0, row - 1, len(headers) - 1)

        workbook.close()
        output.seek(0)

        # File name
        if month:
            file_name = f"stage_summary_report_{year}_{month:02d}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        else:
            file_name = f"stage_summary_report_{year}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"

        attachment = self.env['ir.attachment'].create({
            'name': file_name,
            'type': 'binary',
            'datas': base64.b64encode(output.read()),
            'res_model': 'crm.lead',
            'mimetype': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        })

        return {
            'type': 'ir.actions.act_url',
            'url': f'/web/content/{attachment.id}?download=true',
            'target': 'self',
        }

    def _get_stage_data_by_tracking(self, year, month):
        """
        Build report rows using stage visits that occurred inside the selected month.
        Each stage's lead_count is the number of unique leads that visited that stage during the month.
        """
        # month boundaries
        month_start = date(year, month, 1)
        if month == 12:
            month_end = date(year, 12, 31)
        else:
            month_end = date(year, month + 1, 1) - timedelta(days=1)

        # convert to datetimes for message/tracking date comparisons
        month_start_dt = datetime.combine(month_start, datetime.min.time())
        month_end_dt = datetime.combine(month_end, datetime.max.time())

        # stage_id -> set(lead_ids) for visits during the month
        stage_visits = self._get_stage_visit_counts(date_from=month_start_dt, date_to=month_end_dt)

        result = []
        for stage_id, lead_ids in stage_visits.items():
            # normalize lead_ids to iterable collection
            if isinstance(lead_ids, int):
                lead_ids = [lead_ids]
            elif lead_ids is None:
                lead_ids = []

            stage = self.env['crm.stage'].browse(stage_id)
            # collect sales persons / teams from those leads
            if lead_ids:
                lead_records = self.env['crm.lead'].browse(list(lead_ids))
            else:
                lead_records = self.env['crm.lead']

            sales_persons = {l.user_id.name for l in lead_records if l.user_id}
            sales_teams = {l.team_id.name for l in lead_records if l.team_id}

            result.append({
                'year': year,
                'month': month,
                'month_name': month_start.strftime('%B'),
                'stage_name': stage.name or "Undefined Stage",
                'lead_count': len(lead_ids),
                'sales_persons': ', '.join(sorted(sales_persons)) or 'N/A',
                'sales_teams': ', '.join(sorted(sales_teams)) or 'N/A'
            })

        # sort by lead_count desc then name
        result.sort(key=lambda x: (-x['lead_count'], x['stage_name']))
        return result

    @api.model
    def get_available_years(self):
        leads = self.env['crm.lead'].search([], order='create_date asc', limit=1)
        if not leads:
            return [datetime.now().year]
        start_year = leads.create_date.year
        current_year = datetime.now().year
        return list(range(start_year, current_year + 1))


class CrmStageReportWizard(models.TransientModel):
    _name = 'crm.stage.report.wizard'
    _description = 'CRM Stage Report Wizard'

    year = fields.Selection(
        selection='_get_year_selection',
        string='Year',
        default=lambda self: str(datetime.now().year),
        required=True
    )

    month = fields.Selection([
        ('0', 'All Months'),
        ('1', 'January'),
        ('2', 'February'),
        ('3', 'March'),
        ('4', 'April'),
        ('5', 'May'),
        ('6', 'June'),
        ('7', 'July'),
        ('8', 'August'),
        ('9', 'September'),
        ('10', 'October'),
        ('11', 'November'),
        ('12', 'December'),
    ], string='Month', default='0', required=True)

    @api.model
    def _get_year_selection(self):
        current_year = datetime.now().year
        earliest_lead = self.env['crm.lead'].search([], order='create_date asc', limit=1)
        start_year = earliest_lead.create_date.year if earliest_lead else current_year
        return [(str(year), str(year)) for year in range(start_year, current_year + 1)]

    def generate_report(self):
        year = int(self.year)
        month = int(self.month) if self.month != '0' else None
        return self.env['crm.lead'].action_generate_stage_summary_report(year, month)
