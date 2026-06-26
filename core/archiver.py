"""Archive generated reports by project name and date."""
import os
import shutil
import re
from datetime import datetime
from core.path_utils import ensure_inside_base, sanitize_path_component


def normalize_report_month(report_date):
    """Normalize report date to YYYY-MM for archive folder names."""
    text = str(report_date or '').strip()
    if not text:
        return datetime.now().strftime('%Y-%m')

    match = re.search(r'(\d{4})\D*(\d{1,2})', text)
    if match:
        year = int(match.group(1))
        month = int(match.group(2))
        if 1 <= month <= 12:
            return f'{year:04d}-{month:02d}'

    compact = re.search(r'(\d{4})(\d{2})', text)
    if compact:
        year = int(compact.group(1))
        month = int(compact.group(2))
        if 1 <= month <= 12:
            return f'{year:04d}-{month:02d}'

    return (
        text.replace('年', '-')
        .replace('月', '')
        .replace('日', '')
        .replace('/', '-')
        .replace(' ', '')
        .rstrip('-')
    )


def get_archive_path(base_dir, project_name, report_date=None):
    """
    Build archive path: base_dir/项目名称/日期/
    """
    if report_date is None:
        report_date = datetime.now().strftime('%Y-%m')
    safe_project = sanitize_path_component(project_name)
    safe_date = sanitize_path_component(normalize_report_month(report_date))
    return ensure_inside_base(
        base_dir, os.path.join(base_dir, safe_project, safe_date))


def archive_report(file_path, archive_base, project_name, report_date=None):
    """
    Copy the generated report to the archive directory.
    Returns the archived file path.
    """
    archive_dir = get_archive_path(archive_base, project_name, report_date)
    os.makedirs(archive_dir, exist_ok=True)

    filename = os.path.basename(file_path)
    dest = os.path.join(archive_dir, filename)

    # If file exists, add timestamp suffix
    if os.path.exists(dest):
        name, ext = os.path.splitext(filename)
        timestamp = datetime.now().strftime('%H%M%S')
        dest = os.path.join(archive_dir, f'{name}_{timestamp}{ext}')

    shutil.copy2(file_path, dest)
    return dest


def open_archive_dir(archive_base, project_name=None, report_date=None):
    """Open the archive directory in Windows Explorer."""
    if project_name:
        archive_dir = get_archive_path(archive_base, project_name, report_date)
    else:
        archive_dir = archive_base
    os.makedirs(archive_dir, exist_ok=True)
    os.startfile(archive_dir)
