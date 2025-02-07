import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import Font
import numpy as np

# 设置网页标题
st.title("费用申请单付款状态统计")

# 上传“费用申请单”文件
uploaded_expense_file = st.file_uploader("请上传费用申请单 Excel 文件", type=["xlsx"], key="expense")

# 上传“付款单”文件
uploaded_payment_file = st.file_uploader("请上传付款单 Excel 文件", type=["xlsx"], key="payment")

# 定义处理函数
def process_files(expense_file, payment_file):
    # 读取文件
    费用申请单 = pd.read_excel(expense_file, dtype={'单据编号': str})
    付款单 = pd.read_excel(payment_file, dtype={'源单编号': str})

    # 付款预处理（聚合多笔付款）
    付款汇总 = 付款单.groupby('源单编号', as_index=False).agg(
        累计付款=('本次付款金额', 'sum'),
        最后付款日期=('业务日期', 'max')
    )

    # 多维合并
    合并结果 = pd.merge(
        费用申请单[[
            "单据编号", "申请日期", "申请人", "申请组织",
            "核定金额汇总", "事由", "申请借款", 
            "单据状态"
        ]],
        付款汇总,
        left_on="单据编号",
        right_on="源单编号",
        how="left"
    )

    # 计算付款进度（数值化，修正浮点数精度）
    合并结果['付款进度数值'] = (合并结果['累计付款'] / 合并结果['核定金额汇总']).fillna(0).round(6)

    # 计算付款进度（格式化为百分比字符串）
    合并结果['付款进度'] = (
        合并结果['付款进度数值']
        .apply(lambda x: f"{min(x, 1) * 100:.1f}%")  # 确保最大值不超过 100%
    )

    # 计算付款状态
    合并结果['付款状态'] = pd.cut(
        合并结果['付款进度数值'],
        bins=[-1, 0, 1, float('inf')],
        labels=['未付', '部分支付', '全额支付'],
        right=True
    )

    # 强制修正：付款进度数值等于 1 时，设置为“全额支付”
    合并结果['付款状态'] = np.where(
        合并结果['付款进度数值'] == 1, 
        '全额支付', 
        合并结果['付款状态']
    )

    # 筛选“申请借款”为“是”的单据
    合并结果 = 合并结果[合并结果['申请借款'] == '是']

    # 输出优化
    output_cols = [
        '单据编号', '单据状态','申请日期', '申请人', '核定金额汇总',
        '累计付款', '付款进度', '付款状态', 
        '最后付款日期', '事由'
    ]

    return 合并结果[output_cols]

# 检查是否两个文件都已上传
if uploaded_expense_file is not None and uploaded_payment_file is not None:
    try:
        st.success("文件上传成功！")

        # 显示上传的文件内容（可选）
        st.write("费用申请单预览：")
        expense_df = pd.read_excel(uploaded_expense_file, dtype={'单据编号': str})
        st.dataframe(expense_df.head())

        st.write("付款单预览：")
        payment_df = pd.read_excel(uploaded_payment_file, dtype={'源单编号': str})
        st.dataframe(payment_df.head())

        # 点击按钮执行数据处理
        if st.button("生成付款状态统计文件"):
            # 调用处理函数
            result_df = process_files(uploaded_expense_file, uploaded_payment_file)

            # 将结果保存为 Excel 文件
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                result_df.to_excel(writer, index=False, sheet_name="付款状态统计")
            output.seek(0)  # 将指针移动到文件开头

            # 提供下载按钮
            st.download_button(
                label="下载结果文件",
                data=output,
                file_name="付款状态统计结果.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    except Exception as e:
        st.error(f"文件处理失败：{e}")
else:
    st.info("请同时上传费用申请单和付款单文件以继续")
