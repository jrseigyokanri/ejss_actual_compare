import streamlit as st
import pandas as pd
import base64
import datetime
import plotly.graph_objs as go
import plotly.express as px
import glob
from datetime import datetime

# 現在の日付を取得
today = datetime.today()
formatted_today = today.strftime('%Y-%m-%d')

# Streamlitのページ設定
st.set_page_config(
    page_title="EJSS進捗管理",
    layout="wide",
)

def reset_scroll_position():
    st.write(
        """
        <script>
            setTimeout(function() {
                window.scrollTo(0, 0);
            }, 1);
        </script>
        """,
        unsafe_allow_html=True,
    )


def main_page():
    reset_scroll_position()

    # ファイル名パターンによるExcelファイルの読み込み関数
    def load_excel_file_with_pattern(path, pattern):
        excel_files = glob.glob(f"{path}/*{pattern}*.xlsx")
        if excel_files:
            df = pd.read_excel(excel_files[0], engine='openpyxl', skiprows=2)
            return df
        return None

    # 実績データフィルタリング関数
    def filter_actual_data(df):
        # 勘定科目列が46のみのデータに絞り込み
        df = df[df['勘定科目'] == 46]
        return df

    def get_max_date(df):
        max_date = df['伝票日付'].max()
        if pd.isna(max_date):
            return None
        return pd.to_datetime(str(max_date), format='%Y%m%d')
    
    def zenkaku_num(num):
        zenkaku = {
            0: "０", 1: "１", 2: "２", 3: "３", 4: "４", 5: "５", 6: "６", 7: "７", 8: "８", 9: "９"
        }
        return ''.join([zenkaku[int(d)] for d in str(num)])
    
        


    actual_data = load_excel_file_with_pattern(".", "actual")

    if actual_data is not None:
        actual_data = filter_actual_data(actual_data)

    forecast_data = load_excel_file_with_pattern(".", "EJSS")

    # Get the maximum date from the actual data
    max_date = get_max_date(actual_data) if actual_data is not None else None
    formatted_date = max_date.strftime("%Y年%m月%d日") if max_date is not None else ""

    st.title(f'{formatted_today}')
    st.header('EJSS進捗管理 ')





    if actual_data is not None and forecast_data is not None:

        print(forecast_data)
        # 月の選択
        selected_month = 4
        
        #st.selectbox("実績月を選択してください。", options=list(range(1, 13)))

        # 売上実績データの前処理
        actual_data['伝票日付'] = pd.to_datetime(actual_data['伝票日付'], format='%Y%m%d')
        actual_data = actual_data.set_index('伝票日付')

        # 選択された月の売上実績データ
        selected_month_actual = actual_data[actual_data.index.month == selected_month]

        # 売上実績データの集計
        grouped_actual = selected_month_actual.groupby(['営業担当コード', '営業担当名']).agg({'売上本体金額': 'sum', '仕入本体金額': 'sum'})
        grouped_actual['粗利'] = grouped_actual['売上本体金額'] - grouped_actual['仕入本体金額']
        grouped_actual['粗利率'] = (grouped_actual['粗利'] / grouped_actual['売上本体金額']) * 100

        # 予測データの前処理
        current_month = selected_month
        forecast_data = forecast_data.rename(columns={'担当': '営業担当コード', '担当名': '営業担当名'})


        # 今月の予測データ
        # 既存のコード
        grouped_forecast = forecast_data.groupby(['営業担当コード', '営業担当名']).agg({f'売上金額{str(current_month).translate(str.maketrans("0123456789", "０１２３４５６７８９"))}': 'sum', f'粗利金額{str(current_month).translate(str.maketrans("0123456789", "０１２３４５６７８９"))}': 'sum'})
        # 列名を変更
        grouped_forecast = grouped_forecast.rename(columns={
            f'売上金額{str(current_month).translate(str.maketrans("0123456789", "０１２３４５６７８９"))}': f'EJSS売上{str(current_month).translate(str.maketrans("0123456789", "０１２３４５６７８９"))}',
            f'粗利金額{str(current_month).translate(str.maketrans("0123456789", "０１２３４５６７８９"))}': f'EJSS粗利{str(current_month).translate(str.maketrans("0123456789", "０１２３４５６７８９"))}',
        })
        grouped_forecast['粗利率'] = (grouped_forecast[f'EJSS粗利{str(current_month).translate(str.maketrans("0123456789", "０１２３４５６７８９"))}'] / grouped_forecast[f'EJSS売上{str(current_month).translate(str.maketrans("0123456789", "０１２３４５６７８９"))}']) * 100

    # 今月の売上実績と予測金額
        actual_sum = grouped_actual['売上本体金額'].sum()
        forecast_sum = grouped_forecast[f'EJSS売上{str(current_month).translate(str.maketrans("0123456789", "０１２３４５６７８９"))}'].sum()


        # 達成率と足りない金額
        achievement_rate = (actual_sum / forecast_sum) * 100
        difference = forecast_sum - actual_sum

        # 売上金額と粗利金額の達成率を計算
        grouped_actual['売上達成率'] = (grouped_actual['売上本体金額'] / grouped_forecast[f'EJSS売上{str(current_month).translate(str.maketrans("0123456789", "０１２３４５６７８９"))}']) * 100
        grouped_actual['粗利達成率'] = (grouped_actual['粗利'] / grouped_forecast[f'EJSS粗利{str(current_month).translate(str.maketrans("0123456789", "０１２３４５６７８９"))}']) * 100



       # 営業担当名とコードの一覧を作成し、コードで昇順にソート
        sales_staff_df = actual_data[['営業担当名', '営業担当コード']].drop_duplicates().sort_values('営業担当コード')

        # 担当別選択
        sales_staff = st.selectbox("営業担当を選択してください。", options=['全体'] + sales_staff_df['営業担当名'].tolist())

        if sales_staff != "全体":
            filtered_actual = grouped_actual.loc[grouped_actual.index.get_level_values('営業担当名') == sales_staff]
            filtered_forecast = grouped_forecast.loc[grouped_forecast.index.get_level_values('営業担当名') == sales_staff]
        else:
            filtered_actual = grouped_actual
            filtered_forecast = grouped_forecast


        


        # 売上先別に集計 (売上実績データ)
        grouped_actual_customer = selected_month_actual.groupby(['営業担当コード', '営業担当名', '売上先']).agg({'売上本体金額': 'sum', '仕入本体金額': 'sum'})
        grouped_actual_customer['粗利'] = grouped_actual_customer['売上本体金額'] - grouped_actual_customer['仕入本体金額']
        grouped_actual_customer['粗利率'] = (grouped_actual_customer['粗利'] / grouped_actual_customer['売上本体金額']) * 100

        # 売上先別に集計 (予測データ)
        grouped_forecast_customer = forecast_data.groupby(['営業担当コード', '営業担当名', '売上先']).agg({f'売上金額{zenkaku_num(selected_month)}': 'sum', f'粗利金額{zenkaku_num(selected_month)}': 'sum'})
        # 列名を変更
        grouped_forecast_customer = grouped_forecast_customer.rename(columns={
            f'売上金額{zenkaku_num(selected_month)}': f'EJSS売上{zenkaku_num(selected_month)}',
            f'粗利金額{zenkaku_num(selected_month)}': f'EJSS粗利{zenkaku_num(selected_month)}',
        })
        grouped_forecast_customer['粗利率'] = (grouped_forecast_customer[f'EJSS粗利{zenkaku_num(selected_month)}'] / grouped_forecast_customer[f'EJSS売上{zenkaku_num(selected_month)}']) * 100

        # 4カラム表示
        col1, col2, col3, col4 = st.columns(4)

        # 選択された月の売上実績
        actual_sales = filtered_actual['売上本体金額'].sum()
        col2.metric(f"📝{selected_month}月の売上実績", f"{actual_sales:,.0f}円")

        # 選択された月の売上予測
        sales_forecast = filtered_forecast[f'EJSS売上{zenkaku_num(selected_month)}'].sum()
        col1.metric(f"💰{selected_month}月の売上予測", f"{sales_forecast:,.0f}円")

        # 選択された月の粗利実績
        actual_gross_profit = filtered_actual['粗利'].sum()
        col4.metric(f"📝{selected_month}月の粗利実績", f"{actual_gross_profit:,.0f}円")

        # 選択された月の粗利予測
        gross_profit_forecast = filtered_forecast[f'EJSS粗利{zenkaku_num(selected_month)}'].sum()
        col3.metric(f"💰{selected_month}月の粗利予測", f"{gross_profit_forecast:,.0f}円")

        # 4カラム表示
        col1, col2, col3, col4 = st.columns(4)

        sales_difference = actual_sales - sales_forecast
        sales_achievement = (actual_sales / sales_forecast) * 100

        gross_profit_difference = actual_gross_profit - gross_profit_forecast
        gross_profit_achievement = (actual_gross_profit / gross_profit_forecast) * 100

        col1.metric("📝売上実績と予測の差額", f"{sales_difference:,.0f}円")
        col2.metric("💰売上達成率", f"{sales_achievement:.2f}%")
        col3.metric("📝粗利実績と予測の差額", f"{gross_profit_difference:,.0f}円")
        col4.metric("💰粗利達成率", f"{gross_profit_achievement:.2f}%")


        with st.expander("グラフ表示"):


            # 2列表示
            col1, col2 = st.columns(2)

            with col1:
            # 売上実績と予測の差額
                sales_difference = actual_sales - sales_forecast

                # 円チャートのデータ作成
                sales_difference_pie_data = go.Pie(
                    labels=["Achieved", "Remaining", "Exceeded"],
                    values=[min(actual_sales, sales_forecast), max(sales_forecast - actual_sales, 0), max(actual_sales - sales_forecast, 0)],
                    hole=0.5,
                    textinfo="none",
                    marker=dict(colors=["blue", "lightgray", "green"]),
                )

                # 円チャートのレイアウト設定
                sales_difference_pie_layout = go.Layout(
                    title=f"EJSS売上と実績の差額",
                    annotations=[
                        dict(
                            text=f"{sales_difference:+,.0f}円",
                            x=0.5,
                            y=0.5,
                            showarrow=False,
                            font=dict(size=24),
                        )
                    ],
                )

                # 円チャートの描画
                sales_difference_pie_fig = go.Figure(data=[sales_difference_pie_data], layout=sales_difference_pie_layout)
                st.plotly_chart(sales_difference_pie_fig)

                

            with col2:

                # 粗利実績と予測の差額
                gross_profit_difference = actual_gross_profit - gross_profit_forecast

                # 円チャートのデータ作成
                gross_profit_difference_pie_data = go.Pie(
                    labels=["Achieved", "Remaining", "Exceeded"],
                    values=[min(actual_gross_profit, gross_profit_forecast), max(gross_profit_forecast - actual_gross_profit, 0), max(actual_gross_profit - gross_profit_forecast, 0)],
                    hole=0.5,
                    textinfo="none",
                    marker=dict(colors=["blue", "lightgray", "green"]),
                )

                # 円チャートのレイアウト設定
                gross_profit_difference_pie_layout = go.Layout(
                    title=f"EJSS粗利と実績の差額",
                    annotations=[
                        dict(
                            text=f"{gross_profit_difference:+,.0f}円",
                            x=0.5,
                            y=0.5,
                            showarrow=False,
                            font=dict(size=24),
                        )
                    ],
                )

                # 円チャートの描画
                gross_profit_difference_pie_fig = go.Figure(data=[gross_profit_difference_pie_data], layout=gross_profit_difference_pie_layout)
                st.plotly_chart(gross_profit_difference_pie_fig)

            col1, col2 = st.columns(2)

            # 売上達成率の計算
            with col1:
                achievement_percentage = (actual_sales / sales_forecast) * 100
                remaining_percentage = max(100 - achievement_percentage, 0)

                # 円チャートのデータ作成
                achievement_pie_data = go.Pie(
                    labels=["Achieved", "Remaining", "Exceeded"],
                    values=[min(achievement_percentage, 100), remaining_percentage, max(0, achievement_percentage - 100)],
                    hole=0.5,
                    textinfo="none",
                    marker=dict(colors=["blue", "lightgray", "green"]),
                )

                # 円チャートのレイアウト設定
                achievement_pie_layout = go.Layout(
                    title=f"売上達成率",
                    annotations=[
                        dict(
                            text=f"{achievement_percentage:.2f}%",
                            x=0.5,
                            y=0.5,
                            showarrow=False,
                            font=dict(size=24),
                        )
                    ],
                )

                # 円チャートの描画
                achievement_pie_fig = go.Figure(data=[achievement_pie_data], layout=achievement_pie_layout)
                st.plotly_chart(achievement_pie_fig)

            with col2:
                # 粗利達成率の計算
                gross_profit_achievement_percentage = (actual_gross_profit / gross_profit_forecast) * 100
                remaining_gross_profit_percentage = max(100 - gross_profit_achievement_percentage, 0)

                # 円チャートのデータ作成
                gross_profit_achievement_pie_data = go.Pie(
                    labels=["Achieved", "Remaining", "Exceeded"],
                    values=[min(gross_profit_achievement_percentage, 100), remaining_gross_profit_percentage, max(0, gross_profit_achievement_percentage - 100)],
                    hole=0.5,
                    textinfo="none",
                    marker=dict(colors=["blue", "lightgray", "green"]),
                )

                # 円チャートのレイアウト設定
                gross_profit_achievement_pie_layout = go.Layout(
                    title=f"粗利達成率",
                    annotations=[
                        dict(
                            text=f"{gross_profit_achievement_percentage:.2f}%",
                            x=0.5,
                            y=0.5,
                            showarrow=False,
                            font=dict(size=24),
                        )
                    ],
                )

                # 円チャートの描画
                gross_profit_achievement_pie_fig = go.Figure(data=[gross_profit_achievement_pie_data], layout=gross_profit_achievement_pie_layout)
                st.plotly_chart(gross_profit_achievement_pie_fig)


        

            # 2列表示
            col1, col2 = st.columns(2)

            # 売上実績と予測のグラフデータを作成
            actual_sales_trace = go.Bar(
                y=filtered_actual.index.get_level_values('営業担当名'),
                x=filtered_actual['売上本体金額'],
                name="売上実績",
                marker_color="blue",
                orientation='h'  # 横向きの棒グラフにするために追加
            )

            forecast_sales_trace = go.Bar(
                y=filtered_forecast.index.get_level_values('営業担当名'),
                x=filtered_forecast[f'EJSS売上{zenkaku_num(selected_month)}'],
                name="売上予測",
                marker_color="orange",
                orientation='h'  # 横向きの棒グラフにするために追加
            )

            # グラフレイアウトの設定
            layout = go.Layout(
                title=f"{selected_month}月の売上実績と予測",
                xaxis=dict(title="売上金額 (円)"),
                yaxis=dict(title="営業担当者"),
                barmode="group",
                height=len(filtered_actual) * 50 if sales_staff == "全体" else 300,  # 営業担当が選択された場合には固定の高さに設定
                margin=dict(l=250),  # 左側のマージンを調整
                uniformtext=dict(mode="hide", minsize=8),  # テキストの最小サイズを設定
                bargap=0.15,  # 棒グラフの間隔を調整
                bargroupgap=0.1,  # グループ内の棒グラフの間隔を調整
            )
        

            


            # 粗利実績と予測のグラフデータを作成
            actual_gross_profit_trace = go.Bar(
                y=filtered_actual.index.get_level_values('営業担当名'),
                x=filtered_actual['粗利'],
                name="粗利実績",
                marker_color="blue",
                orientation='h'  # 横向きの棒グラフにするために追加
            )

            forecast_gross_profit_trace = go.Bar(
                y=filtered_forecast.index.get_level_values('営業担当名'),
                x=filtered_forecast[f'EJSS粗利{zenkaku_num(selected_month)}'],
                name="粗利予測",
                marker_color="orange",
                orientation='h'  # 横向きの棒グラフにするために追加
            )

            # グラフレイアウトの設定
            layout_gross_profit = go.Layout(
                title=f"{selected_month}月の粗利実績と予測",
                xaxis=dict(title="粗利金額 (円)"),
                yaxis=dict(title="営業担当者"),
                barmode="group",
                height=len(filtered_actual) * 50 if sales_staff == "全体" else 300,  # 営業担当が選択された場合には固定の高さに設定
                margin=dict(l=250),  # 左側のマージンを調整
                uniformtext=dict(mode="hide", minsize=8),  # テキストの最小サイズを設定
                bargap=0.15,  # 棒グラフの間隔を調整
                bargroupgap=0.1,  # グループ内の棒グラフの間隔を調整
                        )

            with col1:
                # 売上実績と予測のグラフの描画
                fig = go.Figure(data=[actual_sales_trace, forecast_sales_trace], layout=layout)
                st.plotly_chart(fig)

            with col2:
                # 粗利実績と予測のグラフの描画
                fig_gross_profit = go.Figure(data=[actual_gross_profit_trace, forecast_gross_profit_trace], layout=layout_gross_profit)
                st.plotly_chart(fig_gross_profit)

       # 選択された営業担当者に応じてデータフレームをフィルタリング
        if sales_staff != "全体":
            filtered_actual_df = grouped_actual.loc[grouped_actual.index.get_level_values('営業担当名') == sales_staff].reset_index()
            filtered_forecast_df = grouped_forecast.loc[grouped_forecast.index.get_level_values('営業担当名') == sales_staff].reset_index()
        else:
            filtered_actual_df = grouped_actual.reset_index()
            filtered_forecast_df = grouped_forecast.reset_index()


        # Actual DataとForecast Dataを表示
        col1, col2 = st.columns(2)

        with col2:
            st.header('実績')
            st.dataframe(filtered_actual_df)

        with col1:
            st.header('EJSS予測')
            st.dataframe(filtered_forecast_df)

        grouped_actual_by_customer = selected_month_actual.groupby(['売上先', '売上先名', '営業担当コード', '営業担当名']).agg({'売上本体金額': 'sum'})
        grouped_forecast_by_customer = forecast_data.groupby(['売上先', '売上先名', '営業担当コード', '営業担当名']).agg({f'売上金額{zenkaku_num(selected_month)}': 'sum'})
        # 列名を変更
        grouped_forecast_by_customer = grouped_forecast_by_customer.rename(columns={
            f'売上金額{zenkaku_num(selected_month)}': f'EJSS売上{zenkaku_num(selected_month)}',
            f'粗利金額{zenkaku_num(selected_month)}': f'EJSS粗利{zenkaku_num(selected_month)}',
        })

        comparison_df = pd.merge(grouped_actual_by_customer.reset_index(), grouped_forecast_by_customer.reset_index(), on=['売上先', '売上先名', '営業担当コード', '営業担当名'], how='outer').fillna(0)

        comparison_df['差額'] = comparison_df['売上本体金額'] - comparison_df[f'EJSS売上{zenkaku_num(selected_month)}']
        comparison_df['達成率'] = (comparison_df['売上本体金額'] / comparison_df[f'EJSS売上{zenkaku_num(selected_month)}']) * 100

        if sales_staff != "全体":
            comparison_df = comparison_df[comparison_df['営業担当名'] == sales_staff]


        st.header("売上先別売上状況")
        
        comparison_df = comparison_df[['売上先', '売上先名', '営業担当コード', '営業担当名', f'EJSS売上{zenkaku_num(selected_month)}', '売上本体金額', '差額', '達成率']]
        st.dataframe(comparison_df)

        # 実績データで粗利を計算し、売上先毎にグループ化します。
        selected_month_actual['粗利'] = selected_month_actual['売上本体金額'] - selected_month_actual['仕入本体金額']
        grouped_actual_gross_profit_by_customer = selected_month_actual.groupby(['売上先', '売上先名', '営業担当コード', '営業担当名']).agg({'粗利': 'sum'})

        # 予測データを売上先毎にグループ化します。
        grouped_forecast_gross_profit_by_customer = forecast_data.groupby(['売上先', '売上先名', '営業担当コード', '営業担当名']).agg({f'粗利金額{zenkaku_num(selected_month)}': 'sum'})

        # 列名を変更
        grouped_forecast_gross_profit_by_customer = grouped_forecast_gross_profit_by_customer.rename(columns={
            f'粗利金額{zenkaku_num(selected_month)}': f'EJSS粗利{zenkaku_num(selected_month)}',
        })

        # 実績データと予測データをマージします。
        comparison_gross_profit_df = pd.merge(grouped_actual_gross_profit_by_customer.reset_index(), grouped_forecast_gross_profit_by_customer.reset_index(), on=['売上先', '売上先名', '営業担当コード', '営業担当名'], how='outer').fillna(0)

        # 粗利の差額と達成率を計算し、データフレームに追加します。
        comparison_gross_profit_df['粗利差額'] = comparison_gross_profit_df['粗利'] - comparison_gross_profit_df[f'EJSS粗利{zenkaku_num(selected_month)}']
        comparison_gross_profit_df['粗利達成率'] = (comparison_gross_profit_df['粗利'] / comparison_gross_profit_df[f'EJSS粗利{zenkaku_num(selected_month)}']) * 100

        


        #選択された営業担当名に基づいてデータフレームをフィルタリングして表示します。
        if sales_staff != "全体":
            comparison_gross_profit_df = comparison_gross_profit_df[comparison_gross_profit_df['営業担当名'] == sales_staff]

        st.header("売上先別粗利状況")

        comparison_gross_profit_df = comparison_gross_profit_df[['売上先', '売上先名', '営業担当コード', '営業担当名', f'EJSS粗利{zenkaku_num(selected_month)}', '粗利', '粗利差額', '粗利達成率']]

        st.dataframe(comparison_gross_profit_df)


        with st.expander("売上先別比較グラフ"):

                    # 2列表示
            col1, col2 = st.columns(2)

            sorted_sales_df = comparison_df.sort_values(by=f'EJSS売上{zenkaku_num(selected_month)}', ascending=True)
            sorted_gp_df = comparison_gross_profit_df.sort_values(by=f'EJSS粗利{zenkaku_num(selected_month)}', ascending=True)



            # 売上実績と予測のグラフデータを作成
            actual_sales_trace = go.Bar(y=sorted_sales_df['売上先名'], x=sorted_sales_df['売上本体金額'], name="売上実績", marker_color="blue", orientation='h')
            forecast_sales_trace = go.Bar(y=sorted_sales_df['売上先名'], x=sorted_sales_df[f'EJSS売上{zenkaku_num(selected_month)}'], name="売上予測", marker_color="orange", orientation='h')

            # 粗利実績と予測のグラフデータを作成
            actual_gross_profit_trace = go.Bar(y=sorted_gp_df['売上先名'], x=sorted_gp_df['粗利'], name="粗利実績", marker_color="blue", orientation='h')
            forecast_gross_profit_trace = go.Bar(y=sorted_gp_df['売上先名'], x=sorted_gp_df[f'EJSS粗利{zenkaku_num(selected_month)}'], name="粗利予測", marker_color="orange", orientation='h')

            # グラフレイアウトの設定
            layout = go.Layout(
                title=f"{selected_month}月の売上実績と予測",
                xaxis=dict(title="売上金額 (円)"),
                yaxis=dict(title="売上先"),
                barmode="group",
                height=len(sorted_sales_df) * 50,
                margin=dict(l=250),
                uniformtext=dict(mode="hide", minsize=8),
                bargap=0.15,
                bargroupgap=0.1,
            )

            layout_gross_profit = go.Layout(
                title=f"{selected_month}月の粗利実績と予測",
                xaxis=dict(title="粗利金額 (円)"),
                yaxis=dict(title="売上先"),
                barmode="group",
                height=len(sorted_gp_df) * 50,
                margin=dict(l=250),
                uniformtext=dict(mode="hide", minsize=8),
                bargap=0.15,
                bargroupgap=0.1,
            )

            with col1:
                # 売上実績と予測のグラフの描画
                fig = go.Figure(data=[actual_sales_trace, forecast_sales_trace], layout=layout)
                st.plotly_chart(fig)

            with col2:
                # 粗利実績と予測のグラフの描画
                fig_gross_profit = go.Figure(data=[actual_gross_profit_trace, forecast_gross_profit_trace], layout=layout_gross_profit)
                st.plotly_chart(fig_gross_profit)




        

if __name__ == "__main__":
    main_page()
