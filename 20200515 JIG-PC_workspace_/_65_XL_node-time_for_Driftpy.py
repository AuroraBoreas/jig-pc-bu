"""
=======================================================
This script is to generate a node-times for Drift measurement from a give Drift schedule in Excel file.

=Why=
Drift measurement is at least 300 hours time span. But it may be interrupted by certain operation.
In this case, it needs an approach to restore measurement from a breakpoint.

=How=
1.copy node-times from Excel worksheet['Drift raw data ']
2.paster onto the script XL_node_times
3.run this script to get codes
4.copy these codes in text pops out
5.paste into _62_CS2000_SW_Control_Drift.py (look at comment in it)

v0, @ZL, 20191121
=======================================================
"""


import os
from datetime import datetime

XL_node_times = """
2019/11/18 19:13
2019/11/21 15:13
2019/11/22 19:13
2019/11/23 15:13
2019/11/25 15:13
2019/11/27 03:13


"""

li = []
for i, j in enumerate(XL_node_times.strip().split('\n')):
    # print(i, type(i))
    # li.append(i.replace('/','-'))
    if len(j) == (10+1+5):
        j = j.replace('/','-') + ':00'
    j = j.replace('/','-')
    rd = datetime.strptime(j, '%Y-%m-%d %H:%M:%S')
    tt = rd.timetuple()

    if i == 0 or i == 11:
        li.append("s.add_job(task_100IRE, 'date', run_date=datetime({0},{1},{2},{3},{4},{5}))".format(*tt))
        li.append("s.add_job(task_20IRE, 'date', run_date=datetime({0},{1},{2},{3},{4},{5})+timedelta(seconds=n))".format(*tt))
    else:
        li.append("s.add_job(task_100IRE, 'date', run_date=datetime({0},{1},{2},{3},{4},{5}))".format(*tt))

fp = 'drift_node_times.txt'
with open(fp, 'w') as f:
    for i in li:
        f.write(i + '\n')

os.startfile(fp)
