import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import os

DB_FILE = "campaign_database.xlsx"

def safe_division(numerator, denominator):
    return numerator / denominator if denominator else 0

def show_all_campaigns_view(df, location_df):
    st.markdown("## Grand Summary of All Campaigns")
    total_sent = df['Emails Sent'].sum()
    total_delivered = df['Delivered'].sum()
    total_opens = df['Unique Opens'].sum()
    total_clicks = df['Unique Clicks'].sum()
    total_bounces = df['Bounces'].sum()

    col1, col2, col3 = st.columns(3)
    col1.metric("Total Campaigns", len(df))
    col1.metric("Total Emails Sent", f"{total_sent:,}")
    col2.metric("Total Delivered", f"{total_delivered:,}", f"{safe_division(total_delivered, total_sent):.2%}")
    col2.metric("Total Bounces", f"{total_bounces:,}", f"-{safe_division(total_bounces, total_sent):.2%}")
    col3.metric("Total Unique Opens", f"{total_opens:,}", f"{safe_division(total_opens, total_delivered):.2%}")
    col3.metric("Total Unique Clicks", f"{total_clicks:,}", f"{safe_division(total_clicks, total_opens):.2%}")

    st.markdown("<hr>", unsafe_allow_html=True)

    col1, col2 = st.columns(2)
    with col1:
        st.markdown("### Campaign Performance")
        df_melted = df.melt(id_vars=['Campaign'], value_vars=['Emails Sent', 'Delivered', 'Unique Opens', 'Unique Clicks'], var_name='Metric', value_name='Count')
        fig_bar = px.bar(df_melted, x="Campaign", y="Count", color="Metric", barmode="group", text_auto=True,
                         color_discrete_map={"Emails Sent": "#1f77b4", "Delivered": "#2ca02c", "Unique Opens": "#FC3030", "Unique Clicks": "#ff7f0e"})
        st.plotly_chart(fig_bar, use_container_width=True)

    with col2:
        st.markdown("### Engagement Breakdown")
        engagement_df = pd.DataFrame({'Metric': ['Opens', 'Clicks', 'Bounces'], 'Count': [total_opens, total_clicks, total_bounces]})
        fig_donut = px.pie(engagement_df, values='Count', names='Metric', title='Overall Engagement', hole=.4, color_discrete_sequence=["#FC3030", "#ff7f0e", "#2ca02c"])
        st.plotly_chart(fig_donut, use_container_width=True)

    st.markdown("<hr>", unsafe_allow_html=True)

    col3, col4 = st.columns(2)
    with col3:
        st.markdown("### Open Rate by Campaign")
        df['Open Rate'] = (df['Unique Opens'] / df['Delivered']).fillna(0)
        open_rate_df = df.sort_values('Open Rate', ascending=True)
        fig_open_rate = px.bar(open_rate_df, x='Open Rate', y='Campaign', orientation='h', text_auto='.2%')
        fig_open_rate.update_traces(marker_color='#FC3030')
        st.plotly_chart(fig_open_rate, use_container_width=True)

    with col4:
        st.markdown("### Device Usage")
        device_df = df[['Mobile', 'Desktop', 'Tablet']].sum().reset_index()
        device_df.columns = ['Device', 'Percentage']
        if device_df['Percentage'].sum() > 0:
            fig_donut_device = px.pie(device_df, values='Percentage', names='Device', title='Device Usage Distribution', hole=0.4, color_discrete_sequence=["#FC3030", "#ff7f0e", "#2ca02c"])
            st.plotly_chart(fig_donut_device, use_container_width=True)
        else:
            st.info("No device usage data available to display.")

    st.markdown("<hr>", unsafe_allow_html=True)

    col5, col6 = st.columns(2)
    with col5:
        st.markdown("### Geographical Open Rate")
        if not location_df.empty:
            country_opens = location_df.groupby('Country')['Opens'].sum().reset_index()

            fig_scatter_geo = px.scatter_geo(
                country_opens,
                locations="Country",
                locationmode='country names',
                size="Opens",
                hover_name="Country",
                projection="natural earth",
                title="Email Opens by Country",
                color_discrete_sequence=["#FC3030"]
            )
            fig_scatter_geo.update_layout(
                margin=dict(l=0, r=0, t=40, b=0)
            )
            st.plotly_chart(fig_scatter_geo, use_container_width=True)
        else:
            st.info("No location data available to display.")

    with col6:
        st.markdown("### Click-to-Open Rate (CTOR) by Campaign")
        df['CTOR'] = (df['Unique Clicks'] / df['Unique Opens']).fillna(0)
        ctor_df = df.sort_values('CTOR', ascending=True)
        fig_ctor = px.bar(ctor_df, x='CTOR', y='Campaign', orientation='h', text_auto='.2%')
        fig_ctor.update_traces(marker_color='#FC3030')
        st.plotly_chart(fig_ctor, use_container_width=True)

def show_single_campaign_view(df, campaign_name):
    st.markdown(f"## Detailed View: {campaign_name}")
    campaign_data = df[df['Campaign'] == campaign_name].iloc[0]

    # --- Detailed Metrics ---
    col1, col2, col3 = st.columns(3)
    col1.metric("Emails Sent", f"{campaign_data['Emails Sent']:,}")
    col1.metric("Delivered", f"{campaign_data['Delivered']:,}", f"{safe_division(campaign_data['Delivered'], campaign_data['Emails Sent']):.2%}")
    col2.metric("Unique Opens", f"{campaign_data['Unique Opens']:,}", f"{safe_division(campaign_data['Unique Opens'], campaign_data['Delivered']):.2%}")
    col2.metric("Unique Clicks", f"{campaign_data['Unique Clicks']:,}", f"{safe_division(campaign_data['Unique Clicks'], campaign_data['Unique Opens']):.2%}")
    col3.metric("Bounces", f"{campaign_data['Bounces']:,}", f"-{safe_division(campaign_data['Bounces'], campaign_data['Emails Sent']):.2%}")
    col3.metric("Unsubscribes", f"{campaign_data['Unsubscribes']:,}")

    st.markdown("<hr>", unsafe_allow_html=True)

    # --- Funnel and Gauge Charts ---
    col1, col2 = st.columns(2)
    with col1:
        st.markdown("### Email Funnel")
        if campaign_data['Emails Sent'] > 0:
            funnel_data = campaign_data[['Emails Sent', 'Delivered', 'Unique Opens', 'Unique Clicks']]
            fig_funnel = go.Figure(go.Funnel(
                y=funnel_data.index, 
                x=funnel_data.values, 
                textinfo="value+percent initial",
                marker={"color": "#FC3030"}))
            st.plotly_chart(fig_funnel, use_container_width=True)
        elif campaign_data['Delivered'] > 0:
            funnel_data = campaign_data[['Delivered', 'Unique Opens', 'Unique Clicks']]
            fig_funnel = go.Figure(go.Funnel(
                y=funnel_data.index, 
                x=funnel_data.values, 
                textinfo="value+percent initial",
                marker={"color": "#FC3030"}))
            st.plotly_chart(fig_funnel, use_container_width=True)
        else:
            st.info("No email delivery data for this campaign; funnel chart cannot be displayed.")

    with col2:
        st.markdown("### Open Rate Gauge")
        if campaign_data['Delivered'] > 0:
            fig_gauge = go.Figure(go.Indicator(
                mode="gauge+number",
                value=safe_division(campaign_data['Unique Opens'], campaign_data['Delivered']) * 100,
                title={'text': "Open Rate (%)"},
                gauge={
                    'axis': {'range': [None, 100]},
                    'bar': {'color': "#FC3030"},
                }))
            st.plotly_chart(fig_gauge, use_container_width=True)
        else:
            st.info("No emails delivered; open rate cannot be calculated.")

    st.markdown("<hr>", unsafe_allow_html=True)

    # --- New Visualizations for Single Campaign View ---
    col3, col4 = st.columns(2)
    with col3:
        st.markdown("### Bounce Breakdown")
        bounce_types = ["Hard Bounces", "Soft Bounces"]
        bounce_counts = [campaign_data[bt] for bt in bounce_types]
        
        if sum(bounce_counts) > 0:
            bounce_df = pd.DataFrame({'Type': bounce_types, 'Count': bounce_counts})
            fig_bounce_pie = px.pie(bounce_df, values='Count', names='Type', title='Bounce Types', hole=0.4, color_discrete_sequence=["#FC3030", "#ff7f0e"])
            st.plotly_chart(fig_bounce_pie, use_container_width=True)
        else:
            st.info("No bounce data for this campaign.")

    with col4:
        st.markdown("### Device Usage for this Campaign")
        device_data_single = [campaign_data['Mobile'], campaign_data['Desktop'], campaign_data['Tablet']]
        device_labels_single = ['Mobile', 'Desktop', 'Tablet']

        if sum(device_data_single) > 0:
            device_df_single = pd.DataFrame({'Device': device_labels_single, 'Percentage': device_data_single})
            fig_device_single = px.pie(device_df_single, values='Percentage', names='Device', title='Device Usage', hole=0.4, color_discrete_sequence=["#FC3030", "#ff7f0e", "#2ca02c"])
            st.plotly_chart(fig_device_single, use_container_width=True)
        else:
            st.info("No device usage data for this campaign.")

def create_dashboard():
    st.set_page_config(layout="wide", page_title="Campaign Dashboard")
    st.markdown("""<style>...</style>""", unsafe_allow_html=True)
    st.title("Interactive Campaign Dashboard")

    if not os.path.exists(DB_FILE):
        st.error(f"Database file not found: {DB_FILE}. Please run `python create_database.py` first to extract the data.")
        return

    # Load data from Excel
    campaign_df = pd.read_excel(DB_FILE, sheet_name="Campaign_Data")
    location_df = pd.read_excel(DB_FILE, sheet_name="Location_Data")

    # --- Interactive Selector ---
    campaign_list = ["All Campaigns"] + campaign_df['Campaign'].tolist()
    selected_campaign = st.selectbox("Choose a campaign to view:", campaign_list)

    if selected_campaign == "All Campaigns":
        show_all_campaigns_view(campaign_df, location_df)
    else:
        show_single_campaign_view(campaign_df, selected_campaign)

if __name__ == '__main__':
    create_dashboard()
