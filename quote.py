import streamlit as st
import pandas as pd
from datetime import timedelta, datetime
import os
import math
from fpdf import FPDF
import io
# NEW: Import the Word logic from your second file
from file import generate_word_quotation


# --- 1. Define the Mapping (Place this at the top of your file) ---
AIRPORT_MAP = {
    "Kenya": "Nairobi",
    "Tanzania": "Kilimanjaro",
    "Uganda": "Entebbe",
    "Rwanda": "Kigali"
}
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
                travel_start = st.date_input("Start Date", value=datetime(2026, 6, 14), format="DD/MM/YYYY")
                travel_end = st.date_input("End Date", value=datetime(2026, 6, 20), format="DD/MM/YYYY")
                
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
                        # FIX: Use correct column reference loc_df['Room Type']
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
                        diff = num_adults - total_pax_in_rooms
                        if diff > 0:
                            st.error(f"⚠️ Room configuration mismatch. Remaining adults to allot: {diff}")
                        else:
                            st.error(f"⚠️ Room configuration mismatch. Over-allotted by: {abs(diff)} adults.")
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
            
            # --- UPDATED ADDITIONAL CHARGES SECTION ---
            st.subheader("6. Additional Charges")
            if 'extra_items' not in st.session_state:
                st.session_state.extra_items = [{'name': '', 'price': 0.0, 'qty': 1}]

            def add_item_row():
                st.session_state.extra_items.append({'name': '', 'price': 0.0, 'qty': 1})

            def remove_item_row(index):
                st.session_state.extra_items.pop(index)

            for i, item in enumerate(st.session_state.extra_items):
                col1, col2, col3, col4 = st.columns([4, 2, 1, 0.5])
                with col1:
                    item['name'] = st.text_input("Item Name", value=item['name'], key=f"item_name_{i}")
                with col2:
                    item['price'] = st.number_input("Price ($)", value=item['price'], min_value=0.0, step=1.0, key=f"item_price_{i}")
                with col3:
                    item['qty'] = st.number_input("Qty", value=item['qty'], min_value=1, step=1, key=f"item_qty_{i}")
                with col4:
                    st.markdown("<br>", unsafe_allow_html=True) # Align button with inputs
                    if st.button("🗑️", key=f"remove_item_{i}"):
                        remove_item_row(i)
                        st.rerun()

            st.button("➕ Add Item Row", on_click=add_item_row)
            
            # Calculate total extra charges
            extra_charges = sum(item['price'] * item['qty'] for item in st.session_state.extra_items)
            extra_charges_details = [{'Item Name': item['name'], 'Price ($)': item['price'], 'Qty': item['qty']} for item in st.session_state.extra_items if item['name']]

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
                        iti_base_data = []

                        start_airport = AIRPORT_MAP.get(selected_country, "Nairobi")

                        for camp in camp_data:
                            for day_count in range(camp['nights']):
                                # Find the correct rate based on the date
                                a_mask = (camp['df']['Date From'] <= calc_date) & (camp['df']['Date To'] >= calc_date)
                                row = camp['df'][a_mask].iloc[0]
                                
                                s_rate = float(row["Single (Cost Per Person/Per Night)"]) if not pd.isna(row["Single (Cost Per Person/Per Night)"]) else 0
                                d_rate = float(row["Double (Cost Per Person/Per Night)"]) if not pd.isna(row["Double (Cost Per Person/Per Night)"]) else 0
                                t_rate = float(row["Triple (Cost Per Person/Per Night)"]) if not pd.isna(row["Triple (Cost Per Person/Per Night)"]) else 0
                                
                                # Validation for missing rates
                                if (camp['s'] > 0 and s_rate == 0) or (camp['d'] > 0 and d_rate == 0) or (camp['t'] > 0 and t_rate == 0):
                                    st.error(f"❌ Rate not found for room selection at {camp['prop']} on {calc_date.date()}")
                                    st.stop()

                                day_acc_cost = (camp['s'] * s_rate) + (camp['d'] * 2 * d_rate) + (camp['t'] * 3 * t_rate)
                                
                                # Calculate Park Fees
                                p_mask = (df_park['Location'] == camp['loc']) & (df_park['Dates From'] <= calc_date) & (df_park['Dates To'] >= calc_date) & (df_park['Travellers  Category'] == 'Adult')
                                p_rate = float(df_park[p_mask].iloc[0]['Park Fee Per Night Per Person in USD'])
                                
                                acc_total += day_acc_cost
                                park_total += (p_rate * num_adults)
                                
                                # Format the breakdown report
                                room_math = []
                                if camp['s'] > 0: room_math.append(f"{camp['s']}S ({camp['s']} Pax x ${s_rate})")
                                if camp['d'] > 0: room_math.append(f"{camp['d']}D ({camp['d']*2} Pax x ${d_rate})")
                                if camp['t'] > 0: room_math.append(f"{camp['t']}T ({camp['t']*3} Pax x ${t_rate})")
                                
                                acc_report += f"{calc_date.date()} | {camp['prop'][:15]} | {', '.join(room_math)} = ${day_acc_cost:,.2f}\n"
                                park_report += f"{calc_date.date()} | {camp['loc'][:15]} | {num_adults} Pax x ${p_rate} = ${p_rate*num_adults:,.2f}\n"
                                
                                # Determine meal plan: Day 1 = LD, Others = BLD
                                current_meal = "LD" if len(iti_base_data) == 0 else "BLD"
                                
                                # DYNAMIC 'FROM' LOGIC: Use start_airport for the very first day
                                current_from = start_airport if len(iti_base_data) == 0 else iti_base_data[-1]["To"]
                                
                                iti_base_data.append({
                                    "Day": f"Day-{len(iti_base_data)+1}",
                                    "From": current_from,
                                    "To": camp['loc'],
                                    "Activities": "Airport Pickup / Transfer" if len(iti_base_data) == 0 else "Game Drive",
                                    "Accommodation": camp['prop'],
                                    "Meal Plan": current_meal
                                })
                                calc_date += timedelta(days=1)

                        # --- 3. Final Day (Dynamic 'To' Airport) ---
                        iti_base_data.append({
                            "Day": f"Day-{len(iti_base_data)+1}", 
                            "From": iti_base_data[-1]["To"], 
                            "To": start_airport, # DYNAMIC: Returns to the country-specific airport
                            "Activities": "Transfer / Airport Drop", 
                            "Accommodation": "End of Trip", 
                            "Meal Plan": "BL"
                        })

                        total_days = (travel_end - travel_start).days + 1
                        veh_rate = float(df_veh.iloc[0]['Cost in USD/Per Day'])
                        total_veh_cost = veh_rate * total_days * num_vehicles
                        comm_val = float(df_comm.iloc[0]['Commission Per Person (USD)'])
                        total_comm = comm_val * num_adults
                        grand_total = acc_total + park_total + total_veh_cost + total_comm + extra_charges

                        # --- REPLACE THE HIGHLIGHTED SECTION WITH THIS ---

                        # Check if we should hide details (default to False if not set)
                        if not st.session_state.get('hide_details', False):
                            st.subheader("Quotation Breakdown")
                            st.markdown("#### 1. Accommodation")
                            st.code(f"{acc_report}Subtotal: ${acc_total:,.2f}")
                            st.markdown("#### 2. Park Fees")
                            st.code(f"{park_report}Subtotal: ${park_total:,.2f}")
                            st.markdown("#### 3. Vehicle & Commission")
                            veh_math = f"{num_vehicles} Vehicle(s) x {total_days} Days x ${veh_rate:,.2f} = ${total_veh_cost:,.2f}"
                            comm_math = f"{num_adults} Adults x ${comm_val} = ${total_comm:,.2f}"
                            
                            extra_charges_report = ""
                            for item in extra_charges_details:
                                extra_charges_report += f"{item['Item Name']} (${item['Price ($)']:,.2f} x {item['Qty']}) = ${item['Price ($)'] * item['Qty']:,.2f}\n"
                            
                            st.code(f"Vehicle: {veh_math}\nCommission: {comm_math}\nAdditional:\n{extra_charges_report}Total Additional: ${extra_charges:,.2f}")

                            st.markdown(f"""
                                <div class="white-total-box">
                                    <span class="total-title-text">TOTAL TRIP COST: ${grand_total:,.2f}</span>
                                    <span class="per-person-text">COST PER PERSON: ${grand_total/num_adults:,.2f}</span>
                                </div>
                            """, unsafe_allow_html=True)
                        # --- END OF REPLACEMENT ---
                        
                        st.session_state.last_quote = {
                            "total": grand_total, "pp": grand_total/num_adults, "adults": num_adults,
                            "country": selected_country, "iti": iti_base_data,
                            "start": travel_start.strftime("%d/%m/%Y"), "end": travel_end.strftime("%d/%m/%Y"),
                            "pkg": f"{total_days}D/{total_nights}N"
                        }
                        st.session_state.calculation_ready = True

                    except Exception as e:
                        st.error(f"Error during calculation: {e}")

# --- WORD DOCUMENT SECTION ---
if st.session_state.get('calculation_ready'):
    st.divider()
    st.subheader("📋 Itinerary Table")
    st.info("""
    📝 **Note:** You can edit the table below before downloading
    * Check/modify the meal plans
    * Add activities if required
    """)
    
    client_name = st.text_input("Client Name", "Guest", key="client_name_input")
    edited_iti = st.data_editor(st.session_state.last_quote['iti'], key="iti_editor", num_rows="dynamic")
    
    # --- 7. DETAILED ITINERARY UI (EXPANDABLE) ---
    
    
    # Using an expander with a suitcase icon
    with st.expander("🐾 Detailed Itinerary (Optional)", expanded=False):
        st.info("Fill this section to add day-by-day descriptions below the Tariff table.")
        
        if 'detailed_iti' not in st.session_state:
            st.session_state.detailed_iti = [] 

        def add_iti_day():
            new_day_num = len(st.session_state.detailed_iti) + 1
            st.session_state.detailed_iti.append({'day': f"Day {new_day_num}", 'details': ''})

        def remove_iti_day(index):
            st.session_state.detailed_iti.pop(index)

        for i, day_item in enumerate(st.session_state.detailed_iti):
            col1, col2, col3 = st.columns([1, 4, 0.5])
            with col1:
                day_item['day'] = st.text_input("Day label", value=day_item['day'], key=f"det_day_ui_{i}")
            with col2:
                day_item['details'] = st.text_area("Detailed Description", value=day_item['details'], key=f"det_desc_ui_{i}", height=68)
            with col3:
                st.markdown("<br>", unsafe_allow_html=True)
                if st.button("🗑️", key=f"remove_det_ui_btn_{i}"):
                    remove_iti_day(i)
                    st.rerun()

        st.button("➕ Add Detailed Day", on_click=add_iti_day, key="add_det_day_btn_final")
    
    # Filter for Word Doc
    clean_detailed_iti = [d for d in st.session_state.detailed_iti if d['details'].strip()]
    st.divider()

    if st.button("📝 Prepare Word Document", key="prepare_word_final_btn"):
        st.session_state.hide_details = True
        q = st.session_state.last_quote
        
        # --- 1. Existing Accommodation Summary Logic ---
        stay_details = []
        for camp in camp_data:
            rooms = []
            if camp['s'] > 0: rooms.append(f"{camp['s']} Single")
            if camp['d'] > 0: rooms.append(f"{camp['d']} Double")
            if camp['t'] > 0: rooms.append(f"{camp['t']} Triple")
            
            room_str = ", ".join(rooms[:-1]) + " and " + rooms[-1] if len(rooms) > 1 else (rooms[0] if rooms else "")
            stay_details.append(f"{room_str} occupancy rooms at {camp['prop']} for {camp['nights']} night(s)")
        
        full_accommodation_summary = ", ".join(stay_details)

        # --- 2. NEW: Additional Charges Summary Logic ---
        # Get names of items that have a name and a price > 0
        extra_names = [item['name'].strip() for item in st.session_state.extra_items if item['name'].strip() and item['price'] > 0]
        
        extras_str = ""
        if extra_names:
            if len(extra_names) == 1:
                extras_str = extra_names[0]
            else:
                extras_str = ", ".join(extra_names[:-1]) + " and " + extra_names[-1]

        doc_params = {
            "client": client_name,
            "country": q['country'],
            "code": f"{client_name[:3].upper()}_{q['country'][:3].upper()}_{datetime.now().strftime('%d%m%Y')}",
            "pkg": q['pkg'],
            "start": q['start'],
            "end": q['end'],
            "iti": edited_iti,
            "adults": q['adults'],
            "total": q['total'],
            "pp": q['pp'],
            "vehicles": num_vehicles,
            "accommodation_summary": full_accommodation_summary,
            "extras_summary": extras_str, # <--- NEW KEY
            "detailed_iti": clean_detailed_iti
        }
        word_bytes = generate_word_quotation(doc_params)
        st.success("✅ Quotation Generated Successfully!")
        st.download_button("📥 Download Quote", word_bytes, f"Quote_{client_name}.docx", key="download_word_btn_final")