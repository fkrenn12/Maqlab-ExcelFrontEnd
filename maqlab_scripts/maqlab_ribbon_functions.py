from pyxll import xl_app
# import settings_task_pane as settings

import settings_task_pane as demo


def on_initialize_button(control):
    xl = xl_app()
    xl.Selection.Value = "Demo - Hello"


def on_start_button(control):
    # demo.show_tk_ctp()
    demo.show_wx_ctp()
