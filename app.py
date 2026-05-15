from flask import Flask, render_template, jsonify, request
import openpyxl
import json
import os
from datetime import datetime, timedelta, date

app = Flask(__name__)

DATA_DIR = os.path.join(os.path.dirname(__file__), 'data')
EXCEL_FILE = os.path.join(DATA_DIR, 'Annual_Timeline.xlsx')
STATUS_FILE = os.path.join(DATA_DIR, 'tasks_status.json')

CATEGORY_KEYWORDS = {
    'Certificates': ['certificate'],
    'Recovery': ['recovery contest', 'recoverable', 'withdrawal'],
    'Training': ['training', 'quiz'],
    'QA & Reporting': ['scores', 'qa check', 'circulate q'],
    'Passouts': ['passout'],
    'ID Cards': ['pvc', 'card'],
    'Appraisals': ['appraisal'],
    'Admissions': ['admission'],
    'Sprint Management': ['sprint'],
    'Financial': ['financial', 'voucher'],
    'Administration': ['science lab', 'section allocation', 'calendar year'],
}


def assign_category(task_name):
    name = task_name.lower()
    for category, keywords in CATEGORY_KEYWORDS.items():
        if any(kw in name for kw in keywords):
            return category
    return 'General'


def parse_date(val):
    if isinstance(val, datetime):
        return val.date()
    if isinstance(val, str):
        for fmt in ['%m/%d/%y', '%d/%m/%Y', '%Y-%m-%d', '%d/%m/%y']:
            try:
                return datetime.strptime(val, fmt).date()
            except ValueError:
                continue
    return None


def load_tasks():
    tasks = []
    wb = openpyxl.load_workbook(EXCEL_FILE)
    ws = wb['Sheet1']
    for i, row in enumerate(ws.iter_rows(min_row=2, values_only=True)):
        name, desc, task_date, priority = row
        if not name:
            continue
        parsed_date = parse_date(task_date)
        tasks.append({
            'id': i,
            'name': str(name),
            'description': str(desc) if desc else '',
            'date': parsed_date.isoformat() if parsed_date else None,
            'priority': str(priority) if priority else 'Low',
            'category': assign_category(str(name)),
        })
    return tasks


def load_status():
    if os.path.exists(STATUS_FILE):
        with open(STATUS_FILE) as f:
            return json.load(f)
    return {}


def save_status(status):
    os.makedirs(DATA_DIR, exist_ok=True)
    with open(STATUS_FILE, 'w') as f:
        json.dump(status, f, indent=2)


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/api/tasks')
def get_tasks():
    tasks = load_tasks()
    status = load_status()
    today = date.today()
    tomorrow = today + timedelta(days=1)

    for task in tasks:
        tid = str(task['id'])
        task_status = status.get(tid, {})
        task['done'] = task_status.get('done', False)
        task['notes'] = task_status.get('notes', '')
        task['completed_at'] = task_status.get('completed_at', '')

        if task['date']:
            task_date = date.fromisoformat(task['date'])
            task['is_tomorrow'] = task_date == tomorrow
            task['is_today'] = task_date == today
            task['is_overdue'] = task_date < today and not task['done']
        else:
            task['is_tomorrow'] = False
            task['is_today'] = False
            task['is_overdue'] = False

    return jsonify(tasks)


@app.route('/api/tasks/<int:task_id>/complete', methods=['POST'])
def complete_task(task_id):
    status = load_status()
    data = request.get_json() or {}
    status[str(task_id)] = {
        'done': True,
        'notes': data.get('notes', ''),
        'completed_at': datetime.now().isoformat(),
    }
    save_status(status)
    return jsonify({'success': True})


@app.route('/api/tasks/<int:task_id>/uncomplete', methods=['POST'])
def uncomplete_task(task_id):
    status = load_status()
    status[str(task_id)] = {'done': False, 'notes': ''}
    save_status(status)
    return jsonify({'success': True})


if __name__ == '__main__':
    app.run(debug=True, port=5000)
