# Python official installed modules
import os
import time
import sys
import openpyxl #noqa 401 //ruff linter


#################################################
# PROCESS DIAGRAM                               #
#################################################
#                                               #
#                                               #
#     v-------------------------+               #
#     main.py                   |               #
#     |---> supp_paths.py-------+               #
#     |   • pre_VALIDATOR.py----+               #
#     |                                         #
#     v                                         #
#  df_omni_merge_v1p5.py                        #
#     |                                         #
#     v                                         #
#  fom_development_3.py                         #
#     |                                         #
#     v                                         #
#  make_wbxl.py                                 #
#                                               #
#         • blockchain_realloactions_just8s.py  #
#                                               #
#################################################

from supp_paths import df_srcF1, df_srcF2, df_srcF3, df_srcF4, df_srcF6 ,df_srcF7, df_srcF8, df_srcF9, df_srcF10, df_srcF11
parent_dir = os.path.dirname(os.path.dirname(__file__))

# send primary .csv set to parent directory
df_srcF1.to_csv(os.path.join(parent_dir, 'Editorial Garth Group.csv'), index=False)
df_srcF2.to_csv(os.path.join(parent_dir, 'Editorial Luis Group.csv'), index=False)
df_srcF3.to_csv(os.path.join(parent_dir, 'Photo Timesheets.csv'), index=False)
df_srcF4.to_csv(os.path.join(parent_dir, 'Pre-Press.csv'), index=False)
df_srcF6.to_csv(os.path.join(parent_dir, 'Entertainment.csv'), index=False)
df_srcF7.to_csv(os.path.join(parent_dir, 'Prod-Dev-Baseball.csv'), index=False)
df_srcF8.to_csv(os.path.join(parent_dir, 'Prod-Dev-Basketball.csv'), index=False)
df_srcF9.to_csv(os.path.join(parent_dir, 'Prod-Dev-Soccer.csv'), index=False)
df_srcF10.to_csv(os.path.join(parent_dir, 'Prod-Dev-Football.csv'), index=False)
df_srcF11.to_csv(os.path.join(parent_dir, 'Master_Codes_timesheets.csv'), index=False)

print("★ Creating .csv snapshot set. One moment please...")
print('')
time.sleep(0.25)

import df_omni_merge_v1p5 #noqa E402 F401 //ruff linter

def animated_progress_timer(duration):
    bar_length = 20
    for i in range(bar_length + 1):
        progress = "◉" * i + "/" * (bar_length - i)
        percentage = int((i / bar_length) * 100)
        sys.stdout.write(f"\rAlmost done - [{progress}] {percentage}% ")
        sys.stdout.flush()
        time.sleep(duration / bar_length)
    sys.stdout.write("\n")
animated_progress_timer(1)
print('')
print('')
