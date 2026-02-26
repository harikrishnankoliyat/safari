import streamlit as st
import pandas as pd
from datetime import timedelta, datetime
import os
import math

# --- PAGE CONFIGURATION ---
st.set_page_config(page_title="Jaws Africa Safari Planner", layout="wide")

# --- CSS: RESPONSIVE UI ---
st.markdown("""
    <style>
    a.header-anchor { display: none !important; }
    [data-testid="stHeaderActionElements"] { display: none !important; }
    .white-total-box {
        background-color: #FFFFFF !important;
        padding: 40px 20px !important; 
        border-radius: 30px !important;
        border: 4px solid #000000 !important;
        text-align: center !important;
        margin: 30px auto !important;
        width: 95% !important; 
        box-shadow: 0px 20px 50px rgba(0,0,0,0.1);
        display: block !important;
    }
    .total-title-text {
        color: #000000 !important;
        font-family: 'Calibri', sans-serif !important;
        font-weight: 900 !important;
        font-size: 55px !important;
        display: block !important;
        white-space: nowrap !important;
    }
    .per-person-text {
        color: #444444 !important; 
        font-family: 'Calibri', sans-serif !important; 
        font-weight: bold !important;
        font-size: 32px !important;
    }
    .indent { margin-left: 40px; padding-left: 20px; border-left: 2px dotted #000; margin-bottom: 20px; }
    @media only screen and (max-width: 800px) {
        .total-title-text { font-size: 26px !important; white-space: normal !important;}
        .per-person-text { font-size: 18px !important; }
    }
    </style>
""", unsafe_allow_html=True)

# --- UTILITY FUNCTIONS ---
def get_available_countries():
    """Scans directory for .xlsx files and ignores temporary Excel owner files."""
    return sorted([f.replace('.xlsx', '') for f in os.listdir('.') 
                   if f.endswith('.xlsx') and not f.startswith('~$')])

def load_country_data(country_name):
    file_path = f"{country_name}.xlsx"
    if not os.path.exists(file_path): return None
    xls = pd.ExcelFile(file_path)
    df_acc = pd.read_excel(xls, 'Accommodation Cost (Adults)')
    df_park = pd.read_excel(xls, 'Park Fees')
    df_comm = pd.read_excel(xls, 'Jaws Africa Commission')
    df_veh = pd.read_excel(xls, 'Vehicle Cost')
    for df, start, end in [(df_acc, 'Date From', 'Date To'), (df_park, 'Dates From', 'Dates To')]:
        df[start] = pd.to_datetime(df[start])
        df[end] = pd.to_datetime(df[end])
    return df_acc, df_park, df_comm, df_veh

st.title("🦁 Jaws Africa Safari Planner")

available_countries = get_available_countries()
if not available_countries:
    st.error("No Excel files found.")
    st.stop()

selected_country = st.selectbox("1. Select Destination Country", options=available_countries, index=None, placeholder="Select a Country")

if selected_country:
    data = load_country_data(selected_country)
    if data:
        df_acc, df_park, df_comm, df_veh = data

        st.subheader("2. Select Parks/Locations")
        all_parks = sorted(df_acc['Location'].unique().tolist())
        selected_parks = []
        cols = st.columns(3)
        for idx, park in enumerate(all_parks):
            if cols[idx % 3].checkbox(park, key=f"park_{park}"):
                selected_parks.append(park)

        if not selected_parks:
            st.warning("⚠️ Please select at least one park to proceed.")
        else:
            st.divider()
            col_a, col_b = st.columns(2)
            
            with col_a:
                st.subheader("3. Travelers & Dates")
                num_adults = st.number_input("Total Number of Adults", min_value=1, value=2)
                travel_start = st.date_input("Start Date", value=datetime(2026, 6, 14))
                travel_end = st.date_input("End Date", value=datetime(2026, 6, 20))
                
                if travel_end < travel_start:
                    st.error("❌ End date cannot be before start date.")
                    total_nights = 0
                else:
                    total_days = (travel_end - travel_start).days + 1
                    total_nights = max(0, total_days - 1)
                    st.info(f"Trip Duration: {total_days} Days / {total_nights} Nights")

            with col_b:
                st.subheader("4. Vehicle Requirements")
                min_veh = math.ceil(num_adults / 6)
                num_vehicles = st.number_input("Number of Vehicles", min_value=1, value=max(1, min_veh))
                if num_vehicles < min_veh:
                    st.warning(f"⚠️ Minimum {min_veh} vehicles required for {num_adults} adults.")

            # --- ACCOMMODATION PLANNING ---
            st.subheader("5. Accommodation & Room Configuration")
            if 'camps_count' not in st.session_state: st.session_state.camps_count = 1
            
            planned_nights, camp_data = 0, []
            for i in range(st.session_state.camps_count):
                with st.container():
                    st.markdown(f"#### Camp/Lodge #{i+1}")
                    st.markdown('<div class="indent">', unsafe_allow_html=True)
                    c1, c2, c3 = st.columns(3)
                    with c1:
                        loc = st.selectbox("Location", selected_parks, key=f"loc_{i}")
                        loc_df = df_acc[df_acc['Location'] == loc]
                    with c2:
                        acc_type = st.selectbox("Room Type", sorted(loc_df['Room Type'].unique()), key=f"type_{i}")
                        type_df = loc_df[loc_df['Room Type'] == acc_type]
                    with c3:
                        prop = st.selectbox("Property", sorted(type_df['Property'].unique()), key=f"prop_{i}")
                        prop_df = type_df[type_df['Property'] == prop]

                    st.markdown("**Room Quantities:**")
                    rc1, rc2, rc3, rc4 = st.columns(4)
                    with rc1: s_rooms = st.number_input("Single Rooms", min_value=0, step=1, key=f"s_{i}")
                    with rc2: d_rooms = st.number_input("Double Rooms", min_value=0, step=1, key=f"d_{i}")
                    with rc3: t_rooms = st.number_input("Triple Rooms", min_value=0, step=1, key=f"t_{i}")
                    
                    total_pax_in_rooms = (s_rooms * 1) + (d_rooms * 2) + (t_rooms * 3)
                    with rc4:
                        rem_n = total_nights - planned_nights
                        nights = st.number_input("Nights", min_value=1, max_value=max(1, rem_n), value=min(1, rem_n) if rem_n > 0 else 1, key=f"n_{i}")

                    if total_pax_in_rooms != num_adults:
                        st.error(f"⚠️ Room pax ({total_pax_in_rooms}) must match Total Adults ({num_adults})")
                    else:
                        st.success(f"✅ Configuration valid.")

                    planned_nights += nights
                    st.markdown('</div>', unsafe_allow_html=True)
                    camp_data.append({
                        "prop": prop, "loc": loc, "nights": nights, "df": prop_df,
                        "s": s_rooms, "d": d_rooms, "t": t_rooms, "valid": (total_pax_in_rooms == num_adults)
                    })

            if planned_nights < total_nights:
                st.warning(f"⚠️ {total_nights - planned_nights} nights remaining.")
                if st.button("➕ Add More Camps"):
                    st.session_state.camps_count += 1
                    st.rerun()
            elif planned_nights == total_nights:
                st.success("✅ All nights are allotted.")

            st.divider()
            extra_charges = st.number_input("Additional Charges ($)", value=0.0)

            if st.button("🚀 GENERATE CALCULATION", type="primary"):
                all_valid = all([c['valid'] for c in camp_data])
                if not all_valid:
                    st.error("Fix room configurations to match total adults.")
                elif planned_nights != total_nights:
                    st.error(f"Allocate exactly {total_nights} nights.")
                else:
                    try:
                        acc_total, park_total = 0.0, 0.0
                        acc_report, park_report = "", ""
                        calc_date = pd.to_datetime(travel_start)
                        
                        for camp in camp_data:
                            for _ in range(camp['nights']):
                                a_mask = (camp['df']['Date From'] <= calc_date) & (camp['df']['Date To'] >= calc_date)
                                
                                # FETCH RATES
                                s_rate = float(camp['df'][a_mask].iloc[0]["Single (Cost Per Person/Per Night)"]) if camp['s'] > 0 else 0
                                d_rate = float(camp['df'][a_mask].iloc[0]["Double (Cost Per Person/Per Night)"]) if camp['d'] > 0 else 0
                                t_rate = float(camp['df'][a_mask].iloc[0]["Triple (Cost Per Person/Per Night)"]) if camp['t'] > 0 else 0
                                
                                # CALC SUBTOTALS
                                s_sub = camp['s'] * s_rate
                                d_sub = (camp['d'] * 2) * d_rate
                                t_sub = (camp['t'] * 3) * t_rate
                                day_acc_cost = s_sub + d_sub + t_sub
                                
                                # PARK MATCHING
                                p_mask = (df_park['Location'] == camp['loc']) & (df_park['Dates From'] <= calc_date) & \
                                         (df_park['Dates To'] >= calc_date) & (df_park['Travellers  Category'] == 'Adult')
                                p_rate = float(df_park[p_mask].iloc[0]['Park Fee Per Night Per Person in USD'])
                                
                                acc_total += day_acc_cost
                                park_total += (p_rate * num_adults)
                                
                                # BUILD DETAILED BREAKDOWN
                                room_breakdown = []
                                if camp['s'] > 0: room_breakdown.append(f"{camp['s']}S ({camp['s']} Pax x ${s_rate})")
                                if camp['d'] > 0: room_breakdown.append(f"{camp['d']}D ({camp['d']*2} Pax x ${d_rate})")
                                if camp['t'] > 0: room_breakdown.append(f"{camp['t']}T ({camp['t']*3} Pax x ${t_rate})")
                                
                                breakdown_str = ", ".join(room_breakdown)
                                acc_report += f"{calc_date.date()} | {camp['prop'][:10]} | {breakdown_str} = ${day_acc_cost:,.2f}\n"
                                park_report += f"{calc_date.date()} | {camp['loc'][:15]} | {num_adults} Pax x ${p_rate} = ${p_rate*num_adults:,.2f}\n"
                                calc_date += timedelta(days=1)

                        total_days = (travel_end - travel_start).days + 1
                        total_veh_cost = float(df_veh.iloc[0]['Cost in USD/Per Day']) * total_days * num_vehicles
                        comm_val = float(df_comm.iloc[0]['Commission Per Person (USD)'])
                        total_comm = comm_val * num_adults
                        grand_total = acc_total + park_total + total_veh_cost + total_comm + extra_charges

                        st.subheader("Quotation Breakdown")
                        st.markdown("#### 1. Accommodation")
                        st.code(f"{acc_report}Subtotal: ${acc_total:,.2f}")
                        st.markdown("#### 2. Park Fees")
                        st.code(f"{park_report}Subtotal: ${park_total:,.2f}")
                        st.markdown("#### 3. Vehicle & Extras")
                        st.code(f"Vehicle: ${total_veh_cost:,.2f}\nCommission: {num_adults} Adults x ${comm_val} = ${total_comm:,.2f}\nAdditional: ${extra_charges:,.2f}")

                        st.markdown(f"""
                            <div class="white-total-box">
                                <span class="total-title-text">TOTAL TRIP COST: ${grand_total:,.2f}</span>
                                <span class="per-person-text">COST PER PERSON: ${grand_total/num_adults:,.2f}</span>
                            </div>
                        """, unsafe_allow_html=True)
                    except Exception as e:
                        st.error(f"Error: {e}")