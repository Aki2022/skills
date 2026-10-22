"""Five explicit report categories shared by L0 and orchestration."""
CATEGORIES = ('executed','skipped','missing','uncovered','screened')


def empty_report():
    return {key:[] for key in CATEGORIES}


def stage(report,ident,exit_code,targets=None):
    report['executed']=[row for row in report['executed'] if row['id']!=ident]
    report['executed'].append({'id':ident,'exit_code':exit_code,**({'targets':targets} if targets is not None else {})})
    report['skipped']=[row for row in report['skipped'] if row['id']!=ident]


def skip(report,ident,reason):
    if not any(row['id']==ident for row in report['executed']):
        report['skipped']=[row for row in report['skipped'] if row['id']!=ident]
        report['skipped'].append({'id':ident,'reason':reason})
