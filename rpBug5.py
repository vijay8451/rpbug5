import operator
import os
import email.utils
import plotly.plotly as py
import plotly.graph_objs as go
import smtplib
import subprocess
import uuid
import xlrd
from email.mime.text import MIMEText
from fauxfactory import gen_string


authorization='914fd3ee-25c6-42f2-3c00-dte77h21620h'
snaptag='6.6.0-7.0'
to_addr='vijay8451@gmail.com'
mydict = {}
filename = '/var/tmp/{}.xls'.format(uuid.uuid4())
rpUrl = 'reportportal.example.com'


def collect_tests_bug():
    loc = (filename)
    wb = xlrd.open_workbook(loc)
    sheet = wb.sheet_by_index(0)

    for n in range(sheet.nrows):
        row = sheet.row_values(n)
        if row[6] == 'FAILED' or row[6] == 'SKIPPED':

            if 'bugzilla' in row[3]:
                testname = row[3].splitlines()[0]
                bugzilla = row[3].split()[-1].rsplit('=',)[-1]

                try:
                    mydict[bugzilla].append(testname)
                except:
                    mydict[bugzilla] = [testname]



def sort_by_values_len():
    collect_tests_bug()
    dict_len= {key: len(value) for key, value in mydict.items()}
    sorted_key_list = sorted(dict_len.items(), key=operator.itemgetter(1), reverse=True)
    sorted_dict = [{item[0]: mydict[item[0]]} for item in sorted_key_list]
    top_bug = [sorted_dict[bug] for bug in range(len(mydict.keys()))]
    return top_bug


def genrate_graf(sdict):

    bugs = [n.keys() for n in sdict]
    bugs_num = [bugs[n].__iter__().__next__() for n in range(len(bugs))]
    tests_count = [len(sdict[bugs_num.index(n)][n]) for n in bugs_num]
    tests_names = [sdict[bugs_num.index(n)][n] for n in bugs_num]
    compo_name = []
    for count in range(len(tests_names)):
        compo_name.insert(count, [])
        for n in tests_names[count]:
            compo_name[count].append(n.rsplit(':')[0].split("/")[-1])
    uniq_compo = [set(n) for n in compo_name]

    trace0 = go.Bar(
        x=['PB #'+n for n in bugs_num],
        y=tests_count,
        text=uniq_compo,
    )

    data = [trace0]
    layout = go.Layout(
        title='Snap {} Top Bugs'.format(snaptag),
    )

    fig = go.Figure(data=data, layout=layout)
    graf_gen = py.plot(fig, auto_open=False, filename='{}'.format(gen_string('alpha')))

    return [graf_gen, os.path.splitext(graf_gen)[0]+".png"]


def SendEmail():
    sort_dict = sort_by_values_len()
    file_png = genrate_graf(sdict=sort_dict)
    from_mail = 'rpbug5@redhat.com'
    to_mail = to_addr
    html = """\
    <html>
      <head></head>
      <body>
        <p><br>
           Here are the Bugs details ..<br>
           <img src="graph_legend.png" />           
        </p>
        <p>
          For more details <a href="https">link</a>.
        </p>
      </body>
    </html>
    """
    html = html.replace("https", file_png[0])
    html = html.replace("graph_legend.png", file_png[1])
    msg = MIMEText(html, 'html')
    msg['To'] = email.utils.formataddr(('Recipient', to_mail))
    msg['From'] = email.utils.formataddr(('Top Bugs', from_mail))
    msg['Subject'] = 'Snap {} tests failures top Bugs'.format(snaptag)

    smtpObj = smtplib.SMTP('localhost')
    smtpObj.sendmail(from_mail, to_mail, msg=msg.as_string())
    smtpObj.quit()


def fetch():
    """."""
    cmd0 = "curl -X GET --header 'Accept: application/json' --header"
    cmd1 = "'Authorization: bearer {}'".format(authorization)
    cmd2 = "'https://{}/api/v1/SATELLITE6/launch?filter.eq.tags={}'".format(rpUrl, snaptag)
    cmd3 = "-k"
    cmd = cmd0 + " " + cmd1 + " " + cmd2 + " " + cmd3

    try:
        s = subprocess.Popen(cmd, shell=True, stdout=subprocess.PIPE).stdout
        status = s.read()
        launch_id = status.decode().split(',')[3].split(':')[1]
    except Exception as exp:
        print(exp)

    rpcmd0 = "curl -X GET --header 'Accept: application/vnd.ms-excel' --header"
    rpcmd1 = "'Authorization: bearer {}'".format(authorization)
    rpcmd2 = "https://{}/api/v1/SATELLITE6/launch/{}/report?view=xls".\
        format(rpUrl, launch_id.split('"')[1])
    rpcmd3 = '> {}'.format(filename)
    rcmd = rpcmd0 + " " + rpcmd1 + " " + rpcmd2 + " " + cmd3 + " " + rpcmd3

    try:
        subprocess.check_output(rcmd, shell=True)
        SendEmail()
    except Exception as exp:
        print(exp)


fetch()
