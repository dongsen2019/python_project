from Charging_Sys import SysInterface
from Property_Fee import PropertyFee
from PMF import PublicMaintenanceFund
from Cost_Merge import CostMerge
from CustomDialog import DialogMerge, DialogAddr
from Direction_For_User import DirectionDialog
from PyQt5.Qt import *

if __name__ == '__main__':
    import sys
    app = QApplication(sys.argv)

    # 定义项目应收清单地址、收费标准地址、保存地址的三个变量
    xmysqd_addr = ""
    sfbz_addr = ""
    save_addr = ""
    wyf_addr = ""
    wxjj_addr = ""

    # 创建用户界面
    sys_interface = SysInterface()

    # 创建对象
    ppt_fee = PropertyFee()
    pmf = PublicMaintenanceFund()
    cost_merge = CostMerge()

    d = DirectionDialog()   # 系统的使用说明，作为全局对象，且不设置父对象，作为顶层窗口展示

    # 定义导入文件的槽函数
    def data_import(addr):
        # 根据导入的文件名称，将地址赋值给地址变量
        if addr.split('/')[-1][0] == '项':
            global xmysqd_addr
            xmysqd_addr = addr
        elif addr.split('/')[-1][0] == '收':
            global sfbz_addr
            sfbz_addr = addr

    # 定义物业费统计槽函数
    def ppt_cal():
        global wyf_addr
        ppt_fee.read_data(xmysqd_addr)  # 读取数据
        ppt_fee.data_processing()  # 处理数据
        if sys_interface.radio_btn_state() == 1:    # 根据radiobutton的状态选择保存的路径
            tmp_addr = xmysqd_addr.split('/')
            tmp_addr.pop()
            tmp_addr = '/'.join(tmp_addr) + "/wyf.xlsx"
            ppt_fee.write_data(tmp_addr)
            wyf_addr = tmp_addr
            sys_interface.pbar_val_max()
        else:
            ppt_fee.write_data(save_addr)
            wyf_addr = save_addr
            sys_interface.pbar_val_max()

    # 定义日常维修基金统计槽函数
    def pmf_cal():
        global wxjj_addr
        pmf.read_data(xmysqd_addr)  # 读取数据
        pmf.data_processing()  # 处理数据
        if sys_interface.radio_btn_state() == 1:  # 根据radiobutton的状态选择保存的路径
            tmp_addr = xmysqd_addr.split('/')
            tmp_addr.pop()
            tmp_addr = '/'.join(tmp_addr) + "/wxjj.xlsx"
            pmf.write_data(tmp_addr)
            wxjj_addr = tmp_addr
            sys_interface.pbar_val_max()
        else:
            pmf.write_data(save_addr)
            wxjj_addr = save_addr
            sys_interface.pbar_val_max()

    # 定义费用合并槽函数
    def fee_merge():
        try:
            cost_merge.read_data(wyf_addr, wxjj_addr)   # 读取数据
            sys_interface.fee_merge_pbar()  # 加入进度条
            cost_merge.data_processing(sfbz_addr)    # 处理数据
            if sys_interface.radio_btn_state() == 1:  # 根据radiobutton的状态选择保存的路径
                tmp_addr = xmysqd_addr.split('/')
                tmp_addr.pop()
                tmp_addr = '/'.join(tmp_addr) + "/fyhb.xlsx"
                cost_merge.write_data(tmp_addr)
                sys_interface.pbar_val_max()
            else:
                print(save_addr, len(save_addr))
                if len(save_addr) != 0:
                    cost_merge.write_data(save_addr)
                    sys_interface.pbar_val_max()
                else:
                    dlg = DialogAddr()
                    dlg.exec()
        except Exception as e:
            print(e)
            dlg = DialogMerge()
            dlg.exec()

    def directions_for_user():
        d.show()

    # 用户自定义的保存地址
    def to_save_addr(saddr):
        global save_addr
        save_addr = saddr

    # 定义界面类信号
    sys_interface.import_data_signal.connect(data_import)

    sys_interface.btn_ppt_cal_clicked_signal.connect(ppt_cal)

    sys_interface.btn_pmf_cal_clicked_signal.connect(pmf_cal)

    sys_interface.btn_fee_merge_clicked_signal.connect(fee_merge)

    sys_interface.directions_for_use_signal.connect(directions_for_user)

    sys_interface.btn_browse_signal.connect(to_save_addr)

    sys_interface.show()

    sys.exit(app.exec_())

