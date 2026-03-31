import streamlit as st
import pandas as pd
from datetime import timedelta, datetime
import os
import math
import io
import time
# NEW: Import the Word logic from your second file
from file import generate_word_quotation
# NEW: Import Database logic
from database import init_db, save_quote_data, search_quotes, delete_quote

# --- INITIALIZE DATABASE ---
init_db()

# --- 1. Define the Mapping ---
AIRPORT_MAP = {
    "Kenya": "Nairobi",
    "Tanzania": "Kilimanjaro",
    "Uganda": "Entebbe",
    "Rwanda": "Kigali"
}

# --- PAGE CONFIGURATION ---
st.set_page_config(page_title="Jaws Africa Safari Planner", layout="wide")

# --- 2. LOGIN & SESSION TIMEOUT LOGIC ---
def check_timeout():
    if "last_activity" in st.session_state:
        # 15 minutes = 900 seconds
        if time.time() - st.session_state.last_activity > 900:
            for key in list(st.session_state.keys()):
                del st.session_state[key]
            st.warning("⚠️ Session expired due to 15 minutes of inactivity. Please login again.")
            st.stop()
        else:
            st.session_state.last_activity = time.time()

if "logged_in" not in st.session_state:
    st.title("🔐 Jaws Africa Admin Login")
    u_input = st.text_input("Username")
    p_input = st.text_input("Password", type="password")
    if st.button("Login"):
        # MASTER ADMIN: Set is_master to True
        if u_input == "masteradmin" and p_input == "MasterPassword123":
            st.session_state.logged_in = True
            st.session_state.is_master = True 
            st.session_state.last_activity = time.time()
            st.rerun()
        # REGULAR ADMIN: Set is_master to False
        elif u_input == "jawsadmin" and p_input == "Lorkulup":
            st.session_state.logged_in = True
            st.session_state.is_master = False
            st.session_state.last_activity = time.time()
            st.rerun()
        else:
            st.error("❌ Invalid Username or Password")
    st.stop()

check_timeout()

# --- 3. NAVIGATION ---
st.sidebar.title("Menu")

# Define the menu options
menu_options = ["Create Quote", "Search Database", "Logout"]

# Check if the selection is changing to "Create Quote"
if "current_page" not in st.session_state:
    st.session_state.current_page = "Create Quote"

# Use a temporary variable to detect the click
choice = st.sidebar.radio("Navigate", menu_options,label_visibility="collapsed")

# If the user clicks "Create Quote", clear the data to start fresh
if choice == "Create Quote" and st.session_state.current_page != "Create Quote":
    # Keep login/activity but wipe everything else
    keys_to_keep = ['logged_in', 'is_master', 'last_activity']
    for k in list(st.session_state.keys()):
        if k not in keys_to_keep:
            del st.session_state[k]
    st.session_state.current_page = "Create Quote"
    st.rerun()

st.session_state.current_page = choice
app_page = choice

if app_page == "Logout":
    for key in list(st.session_state.keys()):
        del key
    st.rerun()
# --- 4. DATABASE SEARCH PAGE ---
# --- 4. DATABASE SEARCH PAGE ---
# --- 4. DATABASE SEARCH PAGE ---
# --- 4. DATABASE SEARCH PAGE (FINAL FIXED TABLE) ---
# --- 4. DATABASE SEARCH PAGE ---
if app_page == "Search Database":
    st.markdown('<p class="section-header">📂 Quote Database</p>', unsafe_allow_html=True)
    search_query = st.text_input("Search Client Name")
    
    from database import delete_quote, search_quotes
    import json

    db_results = search_quotes(search_query)

    if db_results:
        # 1. Prepare Data
        table_rows = []
        for i, row in enumerate(db_results):
            # Parse config to get the Tour Code
            config_data = json.loads(row[4])
            table_rows.append({
                "Tour Code": config_data.get('code', 'N/A'),
                "Client (Country)": f"{row[1]} ({row[2]})",
                "Date": row[3].split(" ")[0],
                "db_id": row[0],
                "config": row[4]
            })
        
        df = pd.DataFrame(table_rows)

        # 2. Display using st.dataframe (This removes the toolbar)
        # We only show the first three columns to the user
        event = st.dataframe(
            df[["Tour Code", "Client (Country)", "Date"]],
            use_container_width=True,
            hide_index=True,
            on_select="rerun",
            selection_mode="single-row"
        )

        # 3. Handle Actions based on Selection
        if len(event.selection.rows) > 0:
            selected_row_idx = event.selection.rows[0]
            real_data = df.iloc[selected_row_idx]
            
            st.write(f"**Selected:** {real_data['Tour Code']}")
            col1, col2 = st.columns(2)
            
            with col1:
                # Word Generation Logic
                saved_config = json.loads(real_data['config'])
                word_bytes = generate_word_quotation(saved_config)
                st.download_button(
                    label=f"📥 Download Quote: {real_data['Client (Country)']}", 
                    data=word_bytes, 
                    file_name=f"Quote_{real_data['Client (Country)']}.docx",
                    type="primary",
                    use_container_width=True
                )
            
            with col2:
                # Delete Logic (Master Admin Only)
                if st.session_state.get("is_master"):
                    if st.button("🗑️ Delete This Quote", type="secondary", use_container_width=True):
                        delete_quote(real_data['db_id'])
                        st.success("Quote deleted!")
                        st.rerun()
    else:
        st.info("No quotes found.")
    
    st.stop()
# --- 5. MAIN GENERATOR PAGE (YOUR ORIGINAL 508 LINES START HERE) ---

# --- CSS: RESPONSIVE UI ---
# --- CSS: PROFESSIONAL MOBILE-RESPONSIVE UI ---
st.markdown("""
    <style>
    /* HIDE STREAMLIT ANCHORS */
    a.header-anchor, .st-emotion-cache-15z92p2, [data-testid="stHeaderActionElements"] { 
        display: none !important; 
    }

    /* DESKTOP TOTAL BOX */
    .white-total-box {
        background-color: #FFFFFF !important;
        padding: 40px 20px !important; 
        border-radius: 30px !important;
        border: 4px solid #000000 !important;
        text-align: center !important;
        margin: 30px auto !important;
        width: 95% !important; 
        box-shadow: 0px 20px 50px rgba(0,0,0,0.1);
    }
    .total-title-text {
        color: #000000 !important; font-family: 'Calibri', sans-serif !important;
        font-weight: 900 !important; font-size: 55px !important;
    }

    /* MOBILE RESPONSIVE RULES */
    @media only screen and (max-width: 768px) {
        /* Force the 5 columns to stack vertically on mobile */
        [data-testid="stHorizontalBlock"] {
            flex-direction: column !important;
        }
        [data-testid="column"] {
            width: 100% !important;
            margin-bottom: 10px !important;
        }
        /* Make buttons fill the width of the phone */
        .stButton button {
            width: 100% !important;
        }
        /* Scale down the big total price for phone screens */
        .total-title-text { font-size: 28px !important; }
        .white-total-box { padding: 20px 10px !important; }
    }
    </style>
""", unsafe_allow_html=True)
st.markdown("""
    <style>
    /* Custom style for Section Numbers/Titles */
    .section-header {
        font-family: 'Calibri', sans-serif !important;
        font-size: 24px !important;
        font-weight: 700 !important;
        color: #FFFFFF !important;
        margin-top: 20px !important;
        margin-bottom: 10px !important;
    }
    </style>
""", unsafe_allow_html=True)
st.markdown("""
    <style>
    /* Hides the floating toolbar (eye, download, search, etc.) */
    [data-testid="stElementToolbar"] {
        display: none !important;
    }
    
    /* Optional: Hides the 'Edit' pencil icon if you want a read-only look */
    button[title="Edit"] {
        display: none !important;
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
    # NEW: Load Children Policy
    df_child_policy = pd.read_excel(xls, 'Children Rates Policy')
    
    for df, start, end in [(df_acc, 'Date From', 'Date To'), (df_park, 'Dates From', 'Dates To')]:
        df[start] = pd.to_datetime(df[start])
        df[end] = pd.to_datetime(df[end])
    return df_acc, df_park, df_comm, df_veh, df_child_policy

st.title("🦁 Jaws Africa Safari Planner")

available_countries = get_available_countries()
if not available_countries:
    st.error("No Excel files found.")
    st.stop()

st.markdown('<p class="section-header">1. Select Destination Country</p>', unsafe_allow_html=True)
selected_country = st.selectbox(
    "Destination", # This label will be hidden
    options=available_countries, 
    index=None, 
    label_visibility="collapsed" 
)

if selected_country:
    data = load_country_data(selected_country)
    if data:
        df_acc, df_park, df_comm, df_veh, df_child_policy = data

        st.markdown('<p class="section-header">2. Select Parks/Locations</p>', unsafe_allow_html=True)
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
            
            # --- 3. TRAVELERS & DATES ---
            st.markdown('<p class="section-header">3. Travelers, Dates & Vehicles</p>', unsafe_allow_html=True)
            client_name = st.text_input("Client Name", value="Guest")
            
            # Row 1: Dates (Start Left, End Right)
            d_col1, d_col2 = st.columns(2)
            with d_col1:
                travel_start = st.date_input("Start Date", value=datetime(2026, 6, 14), format="DD/MM/YYYY")
            with d_col2:
                travel_end = st.date_input("End Date", value=datetime(2026, 6, 20), format="DD/MM/YYYY")
            
            # Date Validation & Duration Info
            if travel_end < travel_start:
                st.error("❌ End Date cannot be before Start Date. Please select valid dates.")
                st.stop()
            else:
                total_days = (travel_end - travel_start).days + 1
                total_nights = max(0, total_days - 1)
                st.info(f"Trip Duration: {total_days} Days / {total_nights} Nights")

            # Row 2: Adults & Vehicles (Related logic)
            child_seats = 0 
            
            v_col1, v_col2 = st.columns(2)
            with v_col1:
                num_adults = st.number_input("Total Number of Adults", min_value=1, value=2)
                adult_names = [f"Adult {i+1}" for i in range(num_adults)]
            
            # Logic for Children (to get child_seats for the vehicle column)
            st.markdown("---")
            has_children = st.radio("Is there any children in your group?", ["No", "Yes"], index=0, horizontal=True)
            child_data = [] 
            
            if has_children == "Yes":
                with st.expander("🐾 Child Details", expanded=True):
                    num_children = st.number_input("Number of Children", min_value=1, step=1, value=1)
                    st.markdown("---")
                    for i in range(num_children):
                        c_col1, c_col2 = st.columns([1, 3])
                        with c_col1:
                            st.markdown(f"<br>Child {i+1}", unsafe_allow_html=True)
                        with c_col2:
                            age = st.number_input(f"Age", min_value=0, max_value=17, value=5, key=f"c_age_{i}", label_visibility="collapsed")
                        
                        child_data.append({
                            "id": f"Child {i+1} (Age {age})",
                            "age": age,
                            "needs_room": True
                        })
                    st.divider()
                    child_seats = st.number_input(f"How many seats required for {num_children} children?", min_value=0, max_value=num_children, value=num_children)
            else:
                child_seats = 0

            # Now fill the Vehicle column using the calculated seats
            total_veh_pax = num_adults + child_seats
            min_veh_required = math.ceil(total_veh_pax / 6)
            
            with v_col2:
                num_vehicles = st.number_input("Number of Vehicles", min_value=1, value=max(1, min_veh_required))
                if num_vehicles < min_veh_required:
                    st.error(f"⚠️ Minimum {min_veh_required} vehicle(s) required for {total_veh_pax} seated travelers.")
                st.caption(f"Calculated for {total_veh_pax} total travelers requiring seats.")

            # --- 4. ACCOMMODATION PLANNING ---
            st.markdown('<p class="section-header">4. Accommodation & Room Configuration</p>', unsafe_allow_html=True)
            if 'camps_count' not in st.session_state: st.session_state.camps_count = 1
            
            pax_needing_rooms = adult_names + [c["id"] for c in child_data]
            total_required_room_pax = len(pax_needing_rooms)

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

                    st.markdown("**Room Quantities & Pax Assignment:**")
                    rc1, rc2, rc3, rc4 = st.columns(4)
                    with rc1: s_count = st.number_input("Single Rooms", min_value=0, step=1, key=f"s_count_{i}")
                    with rc2: d_count = st.number_input("Double Rooms", min_value=0, step=1, key=f"d_count_{i}")
                    with rc3: t_count = st.number_input("Triple Rooms", min_value=0, step=1, key=f"t_count_{i}")
                    
                    with rc4:
                        rem_n = total_nights - planned_nights
                        nights = st.number_input("Nights", min_value=1, max_value=max(1, rem_n), value=min(1, rem_n) if rem_n > 0 else 1, key=f"n_{i}")

                    assigned_at_this_camp = []
                    room_assignments = {"Single": [], "Double": [], "Triple": []}

                    def create_pax_selector(label, count, max_pax, camp_idx, type_key):
                        for r_idx in range(count):
                            available = [p for p in pax_needing_rooms if p not in assigned_at_this_camp]
                            with st.expander(f"➕ Add Pax: {label} Room {r_idx+1} (Limit {max_pax})"):
                                selected_in_room = []
                                for person in available:
                                    if st.checkbox(person, key=f"chk_{camp_idx}_{type_key}_{r_idx}_{person}"):
                                        selected_in_room.append(person)
                                
                                if len(selected_in_room) > max_pax:
                                    st.error(f"⚠️ Capacity Exceeded: Max {max_pax} allowed for {label}.")
                                
                                assigned_at_this_camp.extend(selected_in_room)
                                room_assignments[label].append(selected_in_room)

                    if s_count > 0: create_pax_selector("Single", s_count, 1, i, "s")
                    if d_count > 0: create_pax_selector("Double", d_count, 2, i, "d")
                    if t_count > 0: create_pax_selector("Triple", t_count, 3, i, "t")

                    if len(assigned_at_this_camp) != total_required_room_pax:
                        diff = total_required_room_pax - len(assigned_at_this_camp)
                        if diff > 0:
                            st.error(f"⚠️ Room assignment incomplete. Remaining people: {diff}")
                        else:
                            st.error(f"⚠️ Assignment mismatch. Over-allotted by: {abs(diff)} people.")
                    else:
                        st.success(f"✅ Configuration valid for {total_required_room_pax} people.")

                    planned_nights += nights
                    st.markdown('</div>', unsafe_allow_html=True)
                    camp_data.append({
                        "prop": prop, "loc": loc, "nights": nights, "df": prop_df,
                        "assignments": room_assignments,
                        "valid": (len(assigned_at_this_camp) == total_required_room_pax)
                    })

            if planned_nights < total_nights:
                st.warning(f"⚠️ {total_nights - planned_nights} nights remaining.")
                if st.button("➕ Add More Camps"):
                    st.session_state.camps_count += 1
                    st.rerun()
            elif planned_nights == total_nights:
                st.success("✅ All nights are allotted.")

                st.divider()
                
                # --- 5. ADDITIONAL CHARGES (UPDATED UI) ---
                st.markdown('<p class="section-header">5. Additional Charges</p>', unsafe_allow_html=True)
                if 'extra_items' not in st.session_state:
                    st.session_state.extra_items = [{'name': '', 'a_price': 0.0, 'c_price': 0.0, 'a_sel': [], 'c_sel': [], 'dyn_c': False, 'dyn_prices': {}}]

                def add_item_row(): st.session_state.extra_items.append({'name': '', 'a_price': 0.0, 'c_price': 0.0, 'a_sel': [], 'c_sel': [], 'dyn_c': False, 'dyn_prices': {}})
                def remove_item_row(index): st.session_state.extra_items.pop(index)

                for i_ex, item in enumerate(st.session_state.extra_items):
                    with st.container():
                        st.markdown(f"**Item {i_ex+1}**")
                        c1, c2, c3 = st.columns([3, 1, 2])
                        item['name'] = c1.text_input("Item Name", value=item['name'], key=f"ex_name_{i_ex}")
                        item['a_price'] = c2.number_input("Adult Price ($)", value=item['a_price'], key=f"ex_ap_{i_ex}")
                        item['a_sel'] = c3.multiselect("Assign Adults", adult_names, default=item['a_sel'], key=f"ex_as_{i_ex}")
                        
                        c4, c5, c6 = st.columns([1, 2, 2])
                        item['dyn_c'] = c4.toggle("Dynamic Child Price", value=item['dyn_c'], key=f"ex_dc_{i_ex}")
                        item['c_sel'] = c5.multiselect("Assign Children", [c['id'] for c in child_data], default=item['c_sel'], key=f"ex_cs_{i_ex}")
                        
                        if item['dyn_c'] and item['c_sel']:
                            for c_id in item['c_sel']:
                                if c_id not in item['dyn_prices']: item['dyn_prices'][c_id] = 0.0
                                item['dyn_prices'][c_id] = st.number_input(f"Price for {c_id}", value=item['dyn_prices'][c_id], key=f"dyn_{i_ex}_{c_id}")
                        else:
                            item['c_price'] = c6.number_input("Flat Child Price ($)", value=item.get('c_price', 0.0), key=f"ex_cp_{i_ex}")
                        
                        if st.button("🗑️ Remove Item", key=f"remove_item_{i_ex}"):
                            remove_item_row(i_ex); st.rerun()
                        st.markdown("---")

                st.button("➕ Add Item Row", on_click=add_item_row)
                
                st.divider()
                # --- 7. CALCULATION ENGINE ---
                if st.button("🚀 GENERATE CALCULATION", type="primary"):
                    all_valid = all([c['valid'] for c in camp_data])
                    
                    if num_vehicles < min_veh_required:
                        st.error(f"❌ Cannot generate: You need at least {min_veh_required} vehicles.")
                    elif not all_valid:
                        st.error("❌ Cannot generate: Some room assignments are missing or invalid.")
                    else:
                        try:
                            # Cost Tracking for Price Table (Round-up logic applied here)
                            indiv_costs = {name: {"acc": 0.0, "park": 0.0, "veh": 0.0, "comm": 0.0, "extra": 0.0, "ff": 1.0} for name in adult_names}
                            for c in child_data: indiv_costs[c['id']] = {"acc": 0.0, "extra": 0.0, "park": 0.0, "veh": 0.0, "comm": 0.0, "ff": 0.0}

                            acc_total, park_total = 0.0, 0.0
                            acc_report, park_report = "", ""
                            calc_date = pd.to_datetime(travel_start)
                            iti_base_data = []
                            start_airport = AIRPORT_MAP.get(selected_country, "Nairobi")

                            for camp in camp_data:
                                for day_count in range(camp['nights']):
                                    # --- Accommodation Engine ---
                                    a_mask = (camp['df']['Date From'] <= calc_date) & (camp['df']['Date To'] >= calc_date)
                                    if not any(a_mask):
                                        st.error(f"❌ Rate not found in Excel for {camp['prop']} on {calc_date.date()}"); st.stop()
                                    
                                    rate_row = camp['df'][a_mask].iloc[0]
                                    day_total_acc_cost = 0.0
                                    detailed_math_parts = []

                                    for r_type, assignments in camp['assignments'].items():
                                        adult_rate_pp = float(rate_row[f"{r_type} (Cost Per Person/Per Night)"])
                                        if len(assignments) > 0 and (pd.isna(adult_rate_pp) or adult_rate_pp == 0):
                                            st.error(f"❌ Rate missing for {r_type} room at {camp['prop']}"); st.stop()

                                        adults_this_type = 0
                                        for room_pax_list in assignments:
                                            # Adults logic
                                            adults_in_room = [p for p in room_pax_list if "Adult" in p]
                                            for a in adults_in_room:
                                                indiv_costs[a]['acc'] += adult_rate_pp
                                                adults_this_type += 1
                                                day_total_acc_cost += adult_rate_pp

                                            # Children logic
                                            children_in_room = [p for p in room_pax_list if "Child" in p]
                                            for person in children_in_room:
                                                c_age = int(person.split("Age ")[1].replace(")", ""))
                                                policy = df_child_policy[(df_child_policy['Property'] == camp['prop']) & (df_child_policy['Age From'] <= c_age) & (df_child_policy['Age To'] >= c_age)]
                                                factor = float(policy.iloc[0]['Form Factor']) if not policy.empty else 1.0
                                                
                                                child_cost = adult_rate_pp * factor
                                                indiv_costs[person]['acc'] += child_cost
                                                indiv_costs[person]['ff'] = max(indiv_costs[person]['ff'], factor) 
                                                day_total_acc_cost += child_cost
                                                
                                                child_id = person.split(" (")[0]
                                                detailed_math_parts.append(f"{child_id} ({factor} * ${adult_rate_pp:,.0f})")

                                        if adults_this_type > 0:
                                            detailed_math_parts.insert(0, f"{adults_this_type} Adult(s) in {r_type} (@ ${adult_rate_pp:,.0f})")

                                    math_string = " + ".join(detailed_math_parts)

                                    # --- Park Fee Logic ---
                                    p_mask_a = (df_park['Location'] == camp['loc']) & (df_park['Dates From'] <= calc_date) & (df_park['Dates To'] >= calc_date) & (df_park['Travellers  Category'] == 'Adult')
                                    p_rate_a = float(df_park[p_mask_a].iloc[0]['Park Fee Per Night Per Person in USD'])
                                    for a in adult_names: indiv_costs[a]['park'] += p_rate_a
                                    day_p_total = p_rate_a * num_adults
                                    
                                    day_child_park_total = 0
                                    if has_children == "Yes":
                                        for c in child_data:
                                            p_mask_c = (df_park['Location'] == camp['loc']) & (df_park['Dates From'] <= calc_date) & (df_park['Dates To'] >= calc_date) & (df_park['Travellers  Category'] == 'Child') & (df_park['Age from'] <= c['age']) & (df_park['Age to'] >= c['age'])
                                            if not df_park[p_mask_c].empty:
                                                pr_val = float(df_park[p_mask_c].iloc[0]['Park Fee Per Night Per Person in USD'])
                                                indiv_costs[c['id']]['park'] += pr_val
                                                day_child_park_total += pr_val

                                    acc_total += day_total_acc_cost
                                    park_total += day_p_total + day_child_park_total
                                    acc_report += f"{calc_date.date()} | {camp['prop'][:12]} | {math_string} = ${day_total_acc_cost:,.2f}\n"
                                    park_report += f"{calc_date.date()} | {camp['loc'][:15]} | Adults: ${day_p_total:,.2f} + Kids: ${day_child_park_total:,.2f} = ${(day_p_total + day_child_park_total):,.2f}\n"

                                    iti_base_data.append({
                                        "Day": f"Day-{len(iti_base_data)+1}", "From": start_airport if not iti_base_data else iti_base_data[-1]["To"], "To": camp['loc'],
                                        "Activities": "Airport Pickup" if len(iti_base_data) == 0 else "Game Drive",
                                        "Accommodation": camp['prop'], "Meal Plan": "BLD"
                                    })
                                    calc_date += timedelta(days=1)

                            iti_base_data.append({"Day": f"Day-{len(iti_base_data)+1}", "From": iti_base_data[-1]["To"], "To": start_airport, "Activities": "Airport Drop", "Accommodation": "End", "Meal Plan": "BL"})

                            # Vehicle Split (Paying pax only)
                            total_days_veh = (travel_end - travel_start).days + 1
                            v_rate = float(df_veh.iloc[0]['Cost in USD/Per Day'])
                            total_v_cost = v_rate * total_days_veh * num_vehicles
                            paying_pax_names = [n for n, d in indiv_costs.items() if d['ff'] > 0]
                            v_per_head = total_v_cost / len(paying_pax_names) if paying_pax_names else 0
                            for p in paying_pax_names: indiv_costs[p]['veh'] = v_per_head

                            # Commission logic (Detailed calculation window)
                            comm_report = ""
                            comm_base = float(df_comm.iloc[0]['Commission Per Person (USD)'])
                            for p_name, data in indiv_costs.items():
                                if "Adult" in p_name:
                                    indiv_costs[p_name]['comm'] = comm_base
                                    comm_report += f"{p_name}: $ {comm_base:,.0f}\n"
                                else:
                                    c_comm = comm_base * data['ff']
                                    indiv_costs[p_name]['comm'] = c_comm
                                    comm_report += f"{p_name}: $ {comm_base:,.0f} * {data['ff']} = $ {c_comm:,.2f}\n"
                            
                            total_comm = sum(d['comm'] for d in indiv_costs.values())

                            # Additional Charges Detailed Breakdown
                            extra_report, total_extra = "", 0.0
                            for item in st.session_state.extra_items:
                                if not item['name']: continue
                                item_adult_total = len(item['a_sel']) * item['a_price']
                                item_child_total = 0
                                for c_id in item['c_sel']:
                                    price = item['dyn_prices'].get(c_id, item['c_price']) if item['dyn_c'] else item['c_price']
                                    indiv_costs[c_id]['extra'] += price
                                    item_child_total += price
                                for a_id in item['a_sel']: indiv_costs[a_id]['extra'] += item['a_price']
                                
                                total_extra += item_adult_total + item_child_total
                                extra_report += f"{item['name']} | Adults({len(item['a_sel'])}) ${item_adult_total:,.0f} + Kids({len(item['c_sel'])}) ${item_child_total:,.2f}\n"

                            # Price Breakdown Table (Rounding logic)
                            price_table_data = []
                            for name, d in indiv_costs.items():
                                # Roundup to next integer as requested
                                total_indiv = math.ceil(d['acc'] + d['park'] + d['veh'] + d['comm'] + d['extra'])
                                price_table_data.append({"Category": name, "Cost": total_indiv})

                            grand_total = sum(r['Cost'] for r in price_table_data)

                            country_part = selected_country[:3].upper()
                            name_part = client_name[:3].replace(" ", "").upper() if 'client_name' in locals() else "GUE"
                            date_part = travel_start.strftime('%d%m%Y')
                            tour_code = f"{country_part}-{name_part}-{date_part}"

                            st.session_state.last_quote = {
                                "total": grand_total, 
                                "pp": grand_total/num_adults if not has_children == "Yes" else 0, # Placeholder, updated in file.py logic
                                "adults": num_adults, 
                                "children_count": len(child_data), 
                                "price_table": price_table_data,
                                "country": selected_country, 
                                "iti": iti_base_data, 
                                "start": travel_start.strftime("%d/%m/%Y"), 
                                "end": travel_end.strftime("%d/%m/%Y"),
                                "pkg": f"{total_days_veh}D/{total_nights}N", 
                                "vehicles": num_vehicles, 
                                "accommodation_summary": "",
                                "code": tour_code
                            }
                            st.session_state.calculation_ready = True
                            
                            st.subheader("Quotation Breakdown")
                            st.markdown("#### 1. Accommodation (Detailed Calculation)")
                            st.code(acc_report)
                            st.markdown("#### 2. Park Fees")
                            st.code(park_report)
                            st.markdown("#### 3. Vehicle Costing")
                            st.code(f"Total: {num_vehicles} Vehicle(s) * {total_days_veh} Days * ${v_rate:,.0f} = ${total_v_cost:,.2f}\nPer Paying Pax: ${v_per_head:,.2f}")
                            st.markdown("#### 4. Commission Breakdown")
                            st.code(f"{comm_report}Total Commission: ${total_comm:,.2f}")
                            st.markdown("#### 5. Additional Charges")
                            st.code(extra_report)
                            st.markdown(f'<div class="white-total-box"><span class="total-title-text">TOTAL TRIP COST: ${grand_total}</span></div>', unsafe_allow_html=True)
                            
                        except Exception as e:
                            st.error(f"Error: {e}")

if st.session_state.get('calculation_ready'):
    st.divider()
    st.subheader("📋 Itinerary Table")
    edited_iti = st.data_editor(st.session_state.last_quote['iti'], num_rows="dynamic")
    
    # --- NEW EDITABLE PRICE TABLE ---
    st.subheader("📋 Detailed Price Table (Editable)")
    edited_price_table = st.data_editor(st.session_state.last_quote['price_table'], num_rows="dynamic")
    include_price_table = st.checkbox("Include Price Table in Word File")

    # --- NEW: DETAILED ITINERARY OPTION ---
    with st.expander("🐾 Detailed Itinerary (Optional)", expanded=False):
        if 'detailed_iti' not in st.session_state: st.session_state.detailed_iti = [] 
        def add_iti_day(): st.session_state.detailed_iti.append({'day': f"Day {len(st.session_state.detailed_iti)+1}", 'details': ''})
        def remove_iti_day(index): st.session_state.detailed_iti.pop(index)
        for i, day_item in enumerate(st.session_state.detailed_iti):
            col1, col2, col3 = st.columns([1, 4, 0.5])
            with col1: day_item['day'] = st.text_input("Label", value=day_item['day'], key=f"det_day_{i}")
            with col2: day_item['details'] = st.text_area("Description", value=day_item['details'], key=f"det_desc_{i}")
            with col3: 
                if st.button("🗑️", key=f"rem_det_{i}"): remove_iti_day(i); st.rerun()
        st.button("➕ Add Day", on_click=add_iti_day)
    
    if st.button("📝 Prepare Word Document"):
        q = st.session_state.last_quote
        q['client'] = client_name
        q['iti'] = edited_iti
        q['price_table'] = edited_price_table if include_price_table else None
        q['detailed_iti'] = [d for d in st.session_state.detailed_iti if d['details'].strip()]
        
        # Word Summary Strings
        stay_details = []
        for camp in camp_data:
            rooms = []
            if camp['assignments']['Single']: rooms.append(f"{len(camp['assignments']['Single'])} Single Room(s)")
            if camp['assignments']['Double']: rooms.append(f"{len(camp['assignments']['Double'])} Double Room(s)")
            if camp['assignments']['Triple']: rooms.append(f"{len(camp['assignments']['Triple'])} Triple Room(s)")
            room_str = ", ".join(rooms[:-1]) + " and " + rooms[-1] if len(rooms) > 1 else (rooms[0] if rooms else "")
            stay_details.append(f"{room_str} at {camp['prop']} for {camp['nights']} night(s)")
        q['accommodation_summary'] = ", ".join(stay_details)
        
        extra_names = [item['name'].strip() for item in st.session_state.extra_items if item['name'].strip()]
        q['extras_summary'] = ", ".join(extra_names[:-1]) + " and " + extra_names[-1] if len(extra_names) > 1 else (extra_names[0] if extra_names else "")
        
        word_bytes = generate_word_quotation(q)
        
        # --- SAVE TO DATABASE ---
        from database import save_quote_data
        # We now save the data (q) instead of the file (word_bytes) to save Railway costs
        # Use the new name and pass the 'q' dictionary instead of 'word_bytes'
        save_quote_data(client_name, selected_country, q)
        
        st.success("✅ Quotation Generated & Saved to Database!")
        st.download_button("📥 Download Quote", word_bytes, f"Quote_{client_name}.docx")

    st.divider()
    if st.button("🔄 Start New Quote (Clear All)"):
        # Keep login/activity but wipe rest
        keys_to_keep = ['logged_in', 'last_activity']
        for k in list(st.session_state.keys()):
            if k not in keys_to_keep:
                del st.session_state[k]
        st.rerun()