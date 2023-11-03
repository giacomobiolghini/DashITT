import streamlit as st
import plotly.express as px
import pandas as pd
from io import BytesIO
from pyxlsb import open_workbook as open_xlsb
import warnings
import xlsxwriter 
import subprocess
import numpy as np
import openpyxl as op
import plotly.figure_factory as ff

st.set_page_config(page_title="Exponento", layout="wide")
st.title("BUSINESS INTELLIGENCE")

def to_excel(df):
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    df.to_excel(writer, index=False, sheet_name='Sheet1')
    workbook = writer.book
    worksheet = writer.sheets['Sheet1']
    format1 = workbook.add_format({'num_format': '0.00'}) 
    worksheet.set_column('A:A', None, format1)  
    writer.close()
    processed_data = output.getvalue()
    return processed_data

df=pd.read_excel('ITT_PROVA.xlsx')

col1, col2 = st.columns((2))

#DIVISIONE PER DATA #############################################################################

startDate = df["Date"].min()
endDate = df["Date"].max()

with col1:
    date1 = st.date_input("Start Date", startDate)

with col2:
    date2 = st.date_input("End date", endDate)

df = df[(df["Date"].dt.date >= date1) & (df["Date"].dt.date <= date2)].copy() 

##################################################################################################

#FILTRI A SINISTRA ################################################################################

st.sidebar.header("Choose your filter: ")

#Filter by Agent name
#agentname = st.sidebar.multiselect("Pick the Agent Name", df["AgentName"].unique())
#if not agentname:
#    df2 = df.copy()
#else:
#    df2 = df[df["AgentName"].isin(agentname)]


#Filter by ConsultantName
supplier = st.sidebar.multiselect("Pick Supplier name", df["SupplierName"].unique())
if not supplier:
    df2 = df.copy()
else:
    df2 = df[df["SupplierName"].isin(supplier)]



#Filter by Agent/Customer Country
location = st.sidebar.multiselect("Pick Customer Country", df2["Agent/Customer Country"].unique())
if not location:
    df3 = df2.copy()
else:
    df3 = df2[df2["Agent/Customer Country"].isin(location)]


total = st.sidebar.multiselect("Pick Number of nights", df3["Total"].unique())
if not total:
    df4 = df3.copy()
else:
    df4 = df3[df3["Total"].isin(total)]

status = st.sidebar.multiselect("Pick the status", df4["Status"].unique())

#Filter by Service
#service = st.sidebar.multiselect("Pick Service", df2["Service"].unique())
#if not service:
#    df5 = df4.copy()
#else:
#    df5 = df4[df4["Service"].isin(service)]

#filter by intersection
if not supplier and not location and not total and not status:
    filtered_df = df
elif not location and not total and not status and supplier:
    filtered_df = df[df["SupplierName"].isin(supplier)]
elif not supplier and not total and not status and location:
    filtered_df = df[df["Agent/Customer Country"].isin(location)]
elif not location and not supplier and not status and total:
    filtered_df = df[df["Total"].isin(total)]
elif not total and not location and not supplier and status:
    filtered_df = df[df["Status"].isin(status)]
elif location and total and supplier and not status:
    filtered_df = df3[df3["Total"].isin(total)& df3["Agent/Customer Country"].isin(location) & df3["SupplierName"].isin(supplier)]
elif location and supplier and status and not total:
    filtered_df = df3[df3["Agent/Customer Country"].isin(location) & df3["SupplierName"].isin(supplier) & df3["Status"].isin(status)]
elif location and total and status and not supplier:
    filtered_df = df3[df3["Agent/Customer Country"].isin(location) & df3["Total"].isin(total) & df3["Status"].isin(status)]
elif supplier and total and status and not location:
    filtered_df = df3[df3["Status"].isin(status) & df3["SupplierName"].isin(supplier) & df3["Total"].isin(total)]
elif location and total and not status and not supplier:
    filtered_df = df2[df2["Agent/Customer Country"].isin(location) & df2["Total"].isin(total)]
elif location and status and not total and not supplier:
    filtered_df = df2[df2["Agent/Customer Country"].isin(location) & df2["Status"].isin(status)]
elif location and supplier and not status and not total:
    filtered_df = df2[df2["Agent/Customer Country"].isin(location) & df2["SupplierName"].isin(supplier)]
elif supplier and status and not location and not total:
    filtered_df = df2[df2["SupplierName"].isin(supplier) & df2["Status"].isin(status)]
elif supplier and total and not status and not consultant:
    filtered_df = df2[df2["SupplierName"].isin(supplier) & df2["Total"].isin(total)]
elif total and status and not location and not supplier:
    filtered_df = df2[df2["Status"].isin(status) & df2["Total"].isin(total)]
else:
    filtered_df= df4[df4["ConsultantName"].isin(consultant) & df4["Agent/Customer Country"].isin(location) & df4["Total"].isin(total)& df4["Status"].isin(status)]



supplier_df = filtered_df.groupby(by = ["SupplierName"], as_index = False)["Retail"].sum()
with col1:
    st.subheader("Revenue by Supplier")
    fig = px.bar(supplier_df, x = "SupplierName", y = "Retail", text= ['${:,.2f}'.format(x) for x in supplier_df["Retail"]], template = "seaborn")
    st.plotly_chart(fig, use_container_width = True, height = 200) 


with col2:
    st.subheader("Revenue by Customer Country")
    fig = px.pie(filtered_df, values = "Retail", names = "Agent/Customer Country", hole = 0.5)
    fig.update_traces(text = filtered_df["Agent/Customer Country"], textposition = "outside")
    st.plotly_chart(fig, use_container_width = True)



cl1, cl2 = st.columns((2))
with cl1:
    with st.expander("SupplierName", expanded=True):
        st.write(supplier_df.style.background_gradient(cmap="Blues"))
        csv = supplier_df.to_csv(index = True).encode('utf-8')
        excel = to_excel(supplier_df)
        st.download_button("Download Data CSV", data =csv, file_name= "Supplier.csv", mime = "text/cvs", help = "Click here to dowmload the data as CSV file")
        st.download_button("Download Data XLSX", data =excel, file_name= "Supplier.xlsx",  help = "Click here to dowmload the data as XLSX file")


location_df = filtered_df.groupby(by = ["Agent/Customer Country"], as_index = False)["Retail"].sum()
with cl2:
        with st.expander("Location data", expanded=True):
            st.write(location_df.style.background_gradient(cmap="Blues"))
            csv = location_df.to_csv(index = True).encode('utf-8')
            excel = to_excel(location_df)
            st.download_button("Download Data CSV", data =csv, file_name= "Location.csv", mime = "text/cvs", help = "Click here to dowmload the data as CSV file")
            st.download_button("Download Data XLSX", data =excel, file_name= "Location.xlsx",  help = "Click here to dowmload the data as XLSX file")

#NIGHTS GRAPH
totalnights_df = filtered_df.groupby(by = ["Total"], as_index = False)["Retail"].sum()
st.subheader("Revenue by total nights")
fig = px.bar(totalnights_df, x = "Total", y = "Retail",  template = "seaborn")
st.plotly_chart(fig, use_container_width = True, height = 200) 


#STATUS
my_list=['FI']

filtered_df["Status FI"] = np.where(filtered_df.Status.isin(my_list), 1, 0)
status_df = filtered_df.groupby(by = ["Status"], as_index = False)
st.subheader("Number of full invoices")
fig = px.bar(filtered_df, x = "Status",template = "seaborn")
st.plotly_chart(fig, use_container_width = True, height = 200) 
############################################################################################################################

#TIME SERIES ANALYSIS
filtered_df["month_year"] = filtered_df["Date"].dt.to_period("M")
st.subheader("Time Series Analysis")

linechart = pd.DataFrame(filtered_df.groupby(filtered_df["month_year"].dt.strftime("%Y : %b"))["Retail"].sum()).reset_index()
fig2 = px.line(linechart, x = "month_year", y="Retail", labels = {"Retail":"Amount"}, height=500, width=1000, template = "gridon")
st.plotly_chart(fig2, use_container_wodth=True)

with st.expander("View Data of TimeSeries", expanded=True):
    st.write(linechart.T.style.background_gradient(cmap="Blue"))
    csv = linechart.to_csv(index = True).encode('utf-8')
    excel = to_excel(linechart)
    st.download_button("Download Data CSV", data =csv, file_name= "Time_series.csv", mime = "text/cvs", help = "Click here to dowmload the data as CSV file")
    st.download_button("Download Data XLSX", data =excel, file_name= "Time-series.xlsx",  help = "Click here to dowmload the data as XLSX file")




###############################################
#DATA SUMMARY AND SCATTERPLOT
import plotly.figure_factory as ff
st.subheader(":point_right : Month wise Consultant revenues summary")
with st.expander("Summary_Table", expanded=True):
    st.markdown("Month wise sub-Category Table")
    filtered_df["month"]=filtered_df["Date"].dt.month_name()
    customer_type_year = pd.pivot_table(data = filtered_df, values ="Retail", index = ["ConsultantName"], columns = "month")
    st.write(customer_type_year.style.background_gradient(cmap="Blues"))

#SCATTER PLOT
data1= px.scatter(filtered_df, x="Retail", y= "Cost")
data1['layout'].update(title="Relationship between Revenue and Cost",
                        titlefont = dict(size=20), xaxis = dict(title = "Retail", titlefont = dict(size=19)),
                        yaxis =dict(title = "Cost", titlefont = dict(size=19)))
st.plotly_chart(data1, use_container_width=True)

data2= px.scatter(filtered_df, x="Retail", y= "Nights")
data2['layout'].update(title="Relationship between Revenue and Hotel nights",
                        titlefont = dict(size=20), xaxis = dict(title = "Retail", titlefont = dict(size=19)),
                        yaxis =dict(title = "Hotel nights", titlefont = dict(size=19)))
st.plotly_chart(data2, use_container_width=True)


dat = pd.read_excel('MAPPAMONDO.xlsx')
dat.columns = ['iso','retail']

fig = px.choropleth(dat, locations="iso",
                    color="retail",
                    hover_name="iso", 
                    color_continuous_scale=px.colors.sequential.Darkmint)
st.plotly_chart(fig, use_container_width = True, height = 200) 
#subprocess.Popen(["activate2.bat"])
