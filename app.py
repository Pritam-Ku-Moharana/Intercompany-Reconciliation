import streamlit as st
import time
import pandas as pd
from io import BytesIO
import base64
import streamlit as st
import streamlit.components.v1 as components

# =========================
# PAGE CONFIG
# =========================
st.set_page_config(
    page_title="Intercompany Reconciliation",
    layout="wide"
)

hide_streamlit_style = """
    <style>
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    </style>
"""
st.markdown(hide_streamlit_style, unsafe_allow_html=True)

# =========================
# CUSTOM CSS (Orange / White / Gray)
# =========================
st.markdown("""
<style>

.block-container {
    padding: 10px 50px 20px 50px; 
    /* Top is 10px, Right is 20px, Bottom is 30px, Left is 40px */
}
/* Main background */
.stApp {
    background-color: transparent;
}

/* Header bar */
.header {
    background-color: #ffffff;
    padding: 10px 20px;
    border-bottom: 3px solid #ff7a00;
    display: flex;
    align-items: center;
}

/* Logo */
.header img {
    height: 70px;
    margin-right: 15px;
}

/* Title */
.header h3 {
    color: #ff7a00;
    font-size: 24px;
    margin: 0;
    text-align: center;
}

/* Upload box */
.upload-box {
    background-color: #ffffff;
    padding: 20px;
    border-radius: 8px;
    border: 1px solid #ddd;
}

/* Buttons */
.stButton > button {
    background-color: #ff7a00;
    color: white;
    border-radius: 6px;
    border: none;
    padding: 10px 18px;
    font-size: 16px;
}

.stButton > button:hover {
    background-color: #e66a00;
}

/* Footer */
.footer {
    margin-top: 40px;
    text-align: center;
    color: gray;
}
#bg-canvas {
  position: fixed;
  top: 0;
  left: 0;
  width: 100vw;
  height: 100vh;
  z-index: -1;
}

</style>
""", unsafe_allow_html=True)
bg_html = """
<div id="bg-canvas"></div>

<script src="https://cdn.jsdelivr.net/npm/three@0.152.2/build/three.min.js"></script>

<script>
  const scene = new THREE.Scene();
  const camera = new THREE.PerspectiveCamera(75, window.innerWidth/window.innerHeight, 1, 1000);
  camera.position.z = 400;

  const renderer = new THREE.WebGLRenderer({ alpha: true });
  renderer.setSize(window.innerWidth, window.innerHeight);
  document.getElementById("bg-canvas").appendChild(renderer.domElement);

  const geometry = new THREE.BufferGeometry();
  const count = 2000;
  const positions = [];

  for (let i = 0; i < count; i++) {
    positions.push(
      (Math.random() - 0.5) * 800,
      (Math.random() - 0.5) * 800,
      (Math.random() - 0.5) * 800
    );
  }

  geometry.setAttribute("position", new THREE.Float32BufferAttribute(positions, 3));

  const material = new THREE.PointsMaterial({
    color: 0xff6e00,
    size: 1.5,
    opacity: 0.8,
    transparent: true
  });

  const particles = new THREE.Points(geometry, material);
  scene.add(particles);

  function animate() {
    requestAnimationFrame(animate);
    particles.rotation.y += 0.0007;
    renderer.render(scene, camera);
  }
  animate();
</script>
"""
components.html(bg_html, height=0)

st.markdown("""
<style>
.stApp {
  background: transparent;
}
.block-container {
  background: rgba(15, 23, 42, 0.85);
  padding: 2rem;
  border-radius: 12px;
}
</style>
""", unsafe_allow_html=True)

# =========================
# HEADER WITH LOGO
# =========================

def get_base64_image(path):
    with open(path, "rb") as f:
        return base64.b64encode(f.read()).decode()

img_base64 = get_base64_image("jsl-logo-guide.png")

st.markdown(f"""
<div style="display:flex; align-items:center;">
    <img src="data:image/png;base64,{img_base64}" width="140">
    <h3 style="margin-left: 100px;">
        Intercompany Reconciliation Processing Portal
    </h3>
</div>
""", unsafe_allow_html=True)


st.write("")  # spacing

# =========================
# FILE UPLOAD SECTION
# =========================
st.markdown("<div class='upload-box'>", unsafe_allow_html=True)

st.subheader("Upload Files")

col1, col2, col3, col4 = st.columns(4)

with col1:
    customer_file = st.file_uploader("Customer Ledger Details File", type=["xlsx", "csv"])

with col2:
    vendor_file = st.file_uploader("Vendor Ledger Details File", type=["xlsx", "csv"])

with col3:
    tds_file = st.file_uploader("TDS File", type=["xlsx", "csv"])

with col4:
    gst_tds_file = st.file_uploader("GST_TDS File", type=["xlsx", "csv"])

st.markdown("</div>", unsafe_allow_html=True)

st.write("")

# =========================
# PROCESS BUTTON
# =========================
if st.button("Process Files"):
    if not all([customer_file, vendor_file, tds_file, gst_tds_file]):
        st.error("Please upload all four files.")
    else:
        
        if customer_file and vendor_file:
            customer = pd.read_excel(customer_file)
            vendor = pd.read_excel(vendor_file)
            tds = pd.read_excel(tds_file)
            gst_tds = pd.read_excel(gst_tds_file)
        # =========================
        # LOADING ANIMATION
        # =========================
        with st.spinner("Processing files, please wait..."):
            # --------------------------------------------------
            # 🔴 YOUR PROCESSING LOGIC GOES HERE
            # --------------------------------------------------
            # Example placeholder (remove later)
            output_loc = BytesIO()
            matched_sheet_columns_name = ["Vendor code", 
                                        "Vendor FI Doc", 
                                        "Vendor posting date", 
                                        "Customer code",
                                        "Customer FI Doc",
                                        "Customer/Vendor Document Type",
                                        "Customer posting date",
                                        "Invoice number",	
                                        "Invoice date",	
                                        "invoice amount-rv",
                                        "Posted amount-re",
                                        "TDS",
                                        "GST TDS",
                                        "Sum",
                                        "Diff",
                                        "Remarks"]
            matched_sheet = pd.DataFrame(columns = matched_sheet_columns_name)
            ###########################################################################################
            not_matched_sheet_columns_name_jargons = ["Vendor code", 
                                        "Vendor FI Doc", 
                                        "Vendor posting date", 
                                        "Customer code",
                                        "Customer FI Doc",
                                        "Customer/Vendor Document Type",
                                        "Customer posting date",
                                        "Invoice number",	
                                        "Invoice date",	
                                        "invoice amount-rv",
                                        "Posted amount-re",
                                        "TDS",
                                        "GST TDS",
                                        "Sum",
                                        "Diff",
                                        "Remarks"]
            not_matched_sheet_df = pd.DataFrame(columns = not_matched_sheet_columns_name_jargons)
            ###########################################################################################

            ########################################################################################################################################################
            ########################################################################################################################################################

            ########################################################################################################################################################
            ########################################################################################################################################################
            # =========================================================
            # 1. REMOVE SA & AB (STORE SEPARATELY)
            # =========================================================
            excluded_ab_sa = ["AB", "SA"]

            customer_ab_sa = customer[customer["Typ"].isin(excluded_ab_sa)]
            vendor_ab_sa   = vendor[vendor["Ty"].isin(excluded_ab_sa)]

            customer_main = customer[~customer["Typ"].isin(excluded_ab_sa)]
            vendor_main   = vendor[~vendor["Ty"].isin(excluded_ab_sa)]

            # =========================================================
            # 2. REMOVE DP (DZ / DA / KZ) (STORE SEPARATELY)
            # =========================================================
            customer_dp = customer_main[customer_main["Typ"] == "DP"]
            vendor_dp   = vendor_main[vendor_main["Ty"].isin(["DZ", "DA", "KZ"])]

            customer_main = customer_main[customer_main["Typ"] != "DP"]
            vendor_main   = vendor_main[~vendor_main["Ty"].isin(["DZ", "DA", "KZ"])]


            ####################################################################################################################################################
            # =========================================================
            # 0. RECORDS WITH NO REFERENCE NUMBER (SEPARATE FIRST)
            # =========================================================
            customer_no_ref = customer_main[
                customer_main["Reference"].isna() | (customer_main["Reference"].astype(str).str.strip() == "")
            ]

            vendor_no_ref = vendor_main[
                vendor_main["Reference"].isna() | (vendor_main["Reference"].astype(str).str.strip() == "")
            ]

            customer_ref = customer_main.drop(customer_no_ref.index)
            vendor_ref   = vendor_main.drop(vendor_no_ref.index)

            ########################################################################################################################################################
            ########################################################################################################################################################
            # =========================================================
            # 3. MATCHING LOGIC (REFERENCE + DOC TYPE)
            # =========================================================
            doc_type_map = {
                "RV": ["RE", "MR", "RT", "KR"],
                "DK": ["DK"]
            }

            valid_pairs = pd.DataFrame(
                [(c, v) for c, vs in doc_type_map.items() for v in vs],
                columns=["Typ", "Ty"]
            )

            ###############################################################
            # ---------------------------------
            # 3.1 CUSTOMER PRESENT IN VENDOR
            # ---------------------------------
            matched_df = (
                customer_main
                .merge(vendor_main, on="Reference", how="inner", suffixes=("_cust", "_vend"))
                .merge(valid_pairs, on=["Typ", "Ty"], how="inner")
            )

            ###############################################################
            # ----------------------------
            # 3.2 CUSTOMER NOT IN VENDOR
            # ----------------------------
            customer_not_in_vendor = (
                customer_main
                .merge(
                    vendor_main.merge(valid_pairs, on="Ty"),
                    on=["Reference", "Typ"],
                    how="left",
                    indicator=True
                )
            )

            customer_not_in_vendor = customer_not_in_vendor[
                customer_not_in_vendor["_merge"] == "left_only"
            ].drop(columns="_merge")

            ###############################################################
            # ----------------------------
            # 3.3 VENDOR NOT IN CUSTOMER
            # ----------------------------
            vendor_not_in_customer = (
                vendor_main
                .merge(
                    customer_main.merge(valid_pairs, on="Typ"),
                    on=["Reference", "Ty"],
                    how="left",
                    indicator=True
                )
            )

            vendor_not_in_customer = vendor_not_in_customer[
                vendor_not_in_customer["_merge"] == "left_only"
            ].drop(columns="_merge")

            ###############################################################
            # ----------------------------
            # VENDOR MATCHED IN CUSTOMER
            # ----------------------------
            vendor_matched_in_customer = (
                vendor_main
                .merge(
                    customer_main.merge(valid_pairs, on="Typ"),
                    on=["Reference", "Ty"],
                    how="inner"
                )
            )



            ####################################################################################################################################################
            ####################################################################################################################################################
            # =========================================================
            # 4. SUMMARY
            # =========================================================
            print("--------------------------SUMMARY--------------------------")

            print("Actual shape of Customer sheet : ",customer.shape)
            print("Actual shape of Vendor sheet : ",vendor.shape,"\n")

            print("Customer AB / SA        :", customer_ab_sa.shape)
            print("Vendor AB / SA          :", vendor_ab_sa.shape,"\n")

            print("Customer DP             :", customer_dp.shape)
            print("Vendor DP               :", vendor_dp.shape,"\n")

            print("Customer matched        :", matched_df.shape)
            print("Vendor matched          :", vendor_matched_in_customer.shape,"\n")

            print("Customer not in Vendor  :", customer_not_in_vendor.shape,"\n")
            print("Vendor not in Customer  :", vendor_not_in_customer.shape,"\n")

            print("Customer no Reference   :", customer_no_ref.shape,"\n")
            print("Vendor no Reference     :", vendor_no_ref.shape,"\n")


            # =========================
            # PREPARE OUTPUT DATAFRAMES
            # =========================
            matched_sheet = pd.DataFrame(columns=[
                "Vendor code","Vendor FI Doc","Vendor posting date",
                "Customer code","Customer FI Doc","Customer/Vendor Document Type",
                "Customer posting date","Invoice number","Invoice date",
                "invoice amount-rv","Posted amount-re","TDS","GST TDS",
                "Sum","Diff","Remarks"
            ])

            rv_dc_s_records = pd.DataFrame(columns=matched_df.columns)

            # =========================
            # MAIN LOOP
            # =========================
            end = len(matched_df)

            for sp in range(1, end + 1):

                dpc = matched_df.iloc[sp-1:sp, :]

                # --------------------------------------------------
                # RV + Vendor D/C = S → STORE & SKIP
                # --------------------------------------------------
                if dpc["Typ"].iloc[0] == "RV" and dpc["D/C_vend"].iloc[0] == "S":
                    rv_dc_s_records = pd.concat([rv_dc_s_records, dpc], ignore_index=True)
                    continue

                # --------------------------------------------------
                # VENDOR DETAILS
                # --------------------------------------------------
                vendor_code = int(dpc["Vendor"].iloc[0])
                vendor_fi_doc = int(dpc["DocumentNo_vend"].iloc[0])
                vendor_posting_date = dpc["Doc--Date"].iloc[0].date().strftime('%Y-%m-%d')

                # --------------------------------------------------
                # CUSTOMER DETAILS
                # --------------------------------------------------
                customer_code = dpc["Customer"].iloc[0]
                customer_fi_doc = int(dpc["DocumentNo_cust"].iloc[0])
                customer_document_typ = dpc["Typ"].iloc[0]
                customer_posting_date = dpc["Pstng Date"].iloc[0].date().strftime("%Y-%m-%d")

                # --------------------------------------------------
                # INVOICE DETAILS
                # --------------------------------------------------
                invoice_number = dpc["Reference"].iloc[0]
                invoice_date = dpc["Doc..Date"].iloc[0].date().strftime("%Y-%m-%d")

                invoice_amount_rv = float(dpc["      Amount in DC"].iloc[0])
                invoice_amount_re = abs(float(dpc["     Amount in LC"].iloc[0]))

                # --------------------------------------------------
                # TDS
                # --------------------------------------------------
                tds["Acc Document No"] = tds["Acc Document No"].astype(str).str.replace(".0","")
                vendor_doc = str(dpc["DocumentNo_vend"].iloc[0])

                matched_tds = vendor_doc in tds["Acc Document No"].values
                tds_rows_matched = float(
                    tds.loc[tds["Acc Document No"] == vendor_doc, "WH Tax Amt"].iloc[0]
                ) if matched_tds else 0

                # --------------------------------------------------
                # GST TDS
                # --------------------------------------------------
                gst_tds["FI Document No"] = gst_tds["FI Document No"].astype(str)

                matched_gst = vendor_doc in gst_tds["FI Document No"].values
                gst_tds_rows_matched = float(
                    gst_tds.loc[gst_tds["FI Document No"] == vendor_doc, "LC Tax Amt.(IGST)"].iloc[0]
                ) if matched_gst else 0

                # --------------------------------------------------
                # AMOUNT CHECK
                # --------------------------------------------------
                sum_1 = invoice_amount_re + tds_rows_matched + gst_tds_rows_matched
                diff = round(invoice_amount_rv - sum_1, 3)
                amount_check = invoice_amount_re - invoice_amount_rv

                # --------------------------------------------------
                # REMARKS
                # --------------------------------------------------
                remarks = (
                    "Matched with " + ", ".join([
                        name for name, ok in {
                            "Doc type": True,
                            "reference": True,
                            "amount": amount_check == 0,
                            "within tolerance": -10 <= diff <= 10,
                            "tds": matched_tds,
                            "gsttds": matched_gst
                        }.items() if ok
                    ])
                )

                # --------------------------------------------------
                # APPEND RESULT
                # --------------------------------------------------
                matched_sheet.loc[len(matched_sheet)] = [
                    vendor_code, vendor_fi_doc, vendor_posting_date,
                    customer_code, customer_fi_doc, customer_document_typ,
                    customer_posting_date, invoice_number, invoice_date,
                    invoice_amount_rv, invoice_amount_re,
                    tds_rows_matched, gst_tds_rows_matched,
                    sum_1, diff, remarks
                ]

            # =========================
            # FINAL SPLIT (TOLERANCE)
            # =========================
            matched_sheet["Diff"] = pd.to_numeric(matched_sheet["Diff"], errors="coerce")

            mask = matched_sheet["Diff"].between(-10, 10)

            matched_df_final = matched_sheet.loc[mask]
            not_matched_df   = matched_sheet.loc[~mask]

            # rv_dc_s_records.to_excel(r"C:\Users\pritam.moharana\Jupyter\Intercompany Reconcilation\more dynamic\rv_dc_s_records.xlsx",index=False)

            # not_matched_df = pd.concat([not_matched_df, rv_dc_s_records], ignore_index=True)
            # =========================
            # WRITE TO EXCEL
            # =========================
            with pd.ExcelWriter(output_loc, engine="xlsxwriter") as writer:
                matched_df_final.to_excel(writer, sheet_name="matched", index=False)
                not_matched_df.to_excel(writer, sheet_name="not matched", index=False)

            # =========================
            # SUMMARY
            # =========================
            print("Matched within tolerance :", len(matched_df_final))
            print("Not matched              :", len(not_matched_df))
            print("RV with Vendor D/C = S   :", len(rv_dc_s_records))

            not_matched_sheet_columns_name_jargons_s = ["Vendor code", 
                                        "Vendor FI Doc", 
                                        "Vendor posting date", 
                                        "Customer code",
                                        "Customer FI Doc",
                                        "Customer/Vendor Document Type",
                                        "Customer posting date",
                                        "Invoice number",	
                                        "Invoice date",	
                                        "invoice amount-rv",
                                        "Posted amount-re",
                                        "TDS",
                                        "GST TDS",
                                        "Sum",
                                        "Diff",
                                        "Remarks"]
            not_matched_sheet_df_for_s = pd.DataFrame(columns = not_matched_sheet_columns_name_jargons_s)
            ###########################################################################################


            # adding the rv_dc_s_records to the not_not_matched_df 
            # =========================
            # MAIN LOOP
            # =========================
            end = len(rv_dc_s_records)

            for dc in range(1, end + 1):
                dpcc = rv_dc_s_records.iloc[dc-1:dc, :]
                # --------------------------------------------------
                # VENDOR DETAILS
                # --------------------------------------------------
                vendor_code = int(dpcc["Vendor"].iloc[0])
                vendor_fi_doc = int(dpcc["DocumentNo_vend"].iloc[0])
                vendor_posting_date = dpcc["Doc--Date"].iloc[0].date().strftime('%Y-%m-%d')

                # --------------------------------------------------
                # CUSTOMER DETAILS
                # --------------------------------------------------
                customer_code = dpcc["Customer"].iloc[0]
                customer_fi_doc = int(dpcc["DocumentNo_cust"].iloc[0])
                customer_document_typ = dpcc["Typ"].iloc[0]
                customer_posting_date = dpcc["Pstng Date"].iloc[0].date().strftime("%Y-%m-%d")

                # --------------------------------------------------
                # INVOICE DETAILS
                # --------------------------------------------------
                invoice_number = dpcc["Reference"].iloc[0]
                invoice_date = dpcc["Doc..Date"].iloc[0].date().strftime("%Y-%m-%d")

                invoice_amount_rv = float(dpcc["      Amount in DC"].iloc[0])
                invoice_amount_re = abs(float(dpcc["     Amount in LC"].iloc[0]))



                # --------------------------------------------------
                # REMARKS
                # --------------------------------------------------
                remarks = (
                    "Matched with " + ", ".join([
                        name for name, ok in {
                            "Doc type": True,
                            "reference": True,
                            "having D/C_cust = S" : True
                        }.items() if ok
                    ])
                )



                # --------------------------------------------------
                # APPEND RESULT
                # --------------------------------------------------
                not_matched_sheet_df_for_s.loc[len(not_matched_sheet_df_for_s),['Vendor code', 
                                                        'Vendor FI Doc', 
                                                        'Vendor posting date', 
                                                        'Customer code',
                                                        'Customer FI Doc', 
                                                        'Customer/Vendor Document Type',
                                                        'Customer posting date', 
                                                        'Invoice number', 
                                                        'Invoice date',
                                                        'invoice amount-rv', 
                                                        'Posted amount-re', 
                                                        'Remarks']] = [vendor_code, 
                                                        vendor_fi_doc, 
                                                        vendor_posting_date, 
                                                        customer_code, 
                                                        customer_fi_doc, 
                                                        customer_document_typ, 
                                                        customer_posting_date, 
                                                        invoice_number, 
                                                        invoice_date, 
                                                        invoice_amount_rv, 
                                                        invoice_amount_re,  
                                                        remarks]
            not_matched_df = pd.concat([not_matched_df, not_matched_sheet_df_for_s],ignore_index = True)


            with pd.ExcelWriter(output_loc, engine="xlsxwriter") as writer:
                matched_df_final.to_excel(writer, sheet_name="matched", index=False)
                not_matched_df.to_excel(writer, sheet_name="not matched", index=False)


            customer_dp["      Amount in DC"] = abs(customer_dp["      Amount in DC"])
            vendor_dp["     Amount in LC"] = abs(vendor_dp["     Amount in LC"])
            customer_dp_sorted = customer_dp.sort_values(by="      Amount in DC")
            vendor_dp_sorted = vendor_dp.sort_values(by = "     Amount in LC")

            ########################################################################################################################################################
            ########################################################################################################################################################
            # records of dp which are present in vice versa

            customer_dp_records_which_are_in_vendor_also = customer_dp_sorted[customer_dp_sorted["      Amount in DC"].isin(
                vendor_dp_sorted["     Amount in LC"]
            )]
            vendor_dp_records_which_are_in_customer_also = vendor_dp_sorted[vendor_dp_sorted["     Amount in LC"].isin(
                customer_dp_records_which_are_in_vendor_also["      Amount in DC"]
            )]

            ########################################################################################################################################################
            ########################################################################################################################################################
            # records of dp which are not present in vice versa
            customer_dp_records_which_are_not_in_vendor = customer_dp_sorted[~customer_dp_sorted["      Amount in DC"].isin(
                vendor_dp_sorted["     Amount in LC"]
            )]

            vendor_dp_records_which_are_not_in_customer = vendor_dp_sorted[~vendor_dp_sorted["     Amount in LC"].isin(
                customer_dp_records_which_are_in_vendor_also["      Amount in DC"]
            )]


            dp_matched_sheet_columns_name = ["Vendor code", 
                                        "Vendor FI Doc", 
                                        "Vendor posting date", 
                                        "Customer code",
                                        "Customer FI Doc",
                                        "Customer/Vendor Document Type",
                                        "Customer posting date",
                                        "Invoice number",	
                                        "Invoice date",	
                                        "invoice amount-rv",
                                        "Posted amount-re",
                                        "TDS",
                                        "GST TDS",
                                        "Sum",
                                        "Diff",
                                        "Remarks"]
            dp_matched_sheet = pd.DataFrame(columns = dp_matched_sheet_columns_name)


            # needed to add the 
            customer_dp_records_which_have_dc_s = customer_dp_records_which_are_in_vendor_also[customer_dp_records_which_are_in_vendor_also["D/C"] == "S"]

            customer_dp_records_which_are_in_vendor_also = customer_dp_records_which_are_in_vendor_also[customer_dp_records_which_are_in_vendor_also["D/C"] != "S"]


            len_of_customer_dp_matched = len(customer_dp_records_which_are_in_vendor_also)
            for ty in range(1, len_of_customer_dp_matched+1):
                cdpr = customer_dp_records_which_are_in_vendor_also.iloc[ty-1:ty,:]
                vdpr = vendor_dp_records_which_are_in_customer_also.iloc[ty-1:ty,:]


                
                # vendor code
                dp_matched_vendor_code = vdpr["Vendor"].iloc[0]
                # vendor fi doc
                dp_matched_vendor_fi_doc = vdpr["DocumentNo"].iloc[0]
                # vendor posting date
                dp_matched_vendor_posting_date = vdpr["Doc--Date"].iloc[0]

                # customer code
                dp_matched_customer_code = cdpr["Customer"].iloc[0]
                # customer fi doc
                dp_matched_customer_fi_doc = cdpr["DocumentNo"].iloc[0]
                # customer/vendor document type
                dp_matched_customer_vendor_document_type = cdpr["Typ"].iloc[0]
                # customer posting date
                dp_matched_customer_posting_date = cdpr["Pstng Date"].iloc[0]
                # invoice number
                dp_matched_invoice_number = cdpr["Reference"].iloc[0]
                # invoice date
                dp_matched_invoice_date = cdpr["Doc..Date"].iloc[0]
                
                # invoice amount-rv
                dp_matched_invoice_amount_rv = cdpr["      Amount in DC"].iloc[0]
                # posted amount-re
                dp_matched_posted_amount_re = vdpr["     Amount in LC"].iloc[0]

                # remarks
                dp_matched_remarks = "DP record matched with amount"


                dp_matched_sheet.loc[len(dp_matched_sheet),["Vendor code", 
                                                            "Vendor FI Doc", 
                                                            "Vendor posting date",
                                                            "Customer code",
                                                            "Customer FI Doc",
                                                            "Customer/Vendor Document Type",
                                                            "Customer posting date",
                                                            "Invoice number",
                                                            "Invoice date",
                                                            "invoice amount-rv",
                                                            "Posted amount-re",
                                                            "Remarks"]] = [dp_matched_vendor_code,
                                                                        dp_matched_vendor_fi_doc,
                                                                        dp_matched_vendor_posting_date,
                                                                        dp_matched_customer_code,
                                                                        dp_matched_customer_fi_doc,
                                                                        dp_matched_customer_vendor_document_type,
                                                                        dp_matched_customer_posting_date,
                                                                        dp_matched_invoice_number,
                                                                        dp_matched_invoice_date,
                                                                        dp_matched_invoice_amount_rv,
                                                                        dp_matched_posted_amount_re,
                                                                        dp_matched_remarks]

            matched_df_final = pd.concat([matched_df_final,dp_matched_sheet],ignore_index = True)
            # =========================
            # WRITE TO EXCEL
            # =========================
            with pd.ExcelWriter(output_loc, engine="xlsxwriter") as writer:
                matched_df_final.to_excel(writer, sheet_name="matched", index=False)
                not_matched_df.to_excel(writer, sheet_name="not matched", index=False)


            # insertion of Customer DP Record, which has D/C = S
            len_dp_s = len(customer_dp_records_which_have_dc_s)
            for ok in range(1,len_dp_s+1):
                dpc = customer_dp_records_which_have_dc_s.iloc[ok-1:ok, :]


                # --------------------------------------------------
                # CUSTOMER DETAILS
                # --------------------------------------------------
                customer_code = dpc["Customer"].iloc[0]
                customer_fi_doc = int(dpc["DocumentNo"].iloc[0])
                customer_document_typ = dpc["Typ"].iloc[0]
                customer_posting_date = dpc["Pstng Date"].iloc[0].date().strftime("%Y-%m-%d")

                # --------------------------------------------------
                # INVOICE DETAILS
                # --------------------------------------------------
                invoice_number = dpc["Reference"].iloc[0]
                invoice_date = dpc["Doc..Date"].iloc[0].date().strftime("%Y-%m-%d")
                invoice_amount_rv = float(dpc["      Amount in DC"].iloc[0])

                remarks = "Customer DP Record, which has D/C = S"

                not_matched_df.loc[len(not_matched_df),['Customer code',
                                                        'Customer FI Doc', 
                                                        'Customer/Vendor Document Type',
                                                        'Customer posting date',
                                                        'Invoice number', 
                                                        'Invoice date',
                                                        'invoice amount-rv',
                                                        'Remarks']] = [customer_code,
                                                                    customer_fi_doc,
                                                                    customer_document_typ,
                                                                    customer_posting_date,
                                                                    invoice_number,
                                                                    invoice_date,
                                                                    invoice_amount_rv,
                                                                    remarks]


            # customer dp type records which are not present in the vendor sheet
            len_of_customer_dp_not_present = len(customer_dp_records_which_are_not_in_vendor)
            for bb in range(1,len_of_customer_dp_not_present+1):
                dpc = customer_dp_records_which_are_not_in_vendor.iloc[bb-1:bb, :]


                # --------------------------------------------------
                # CUSTOMER DETAILS
                # --------------------------------------------------
                customer_code = dpc["Customer"].iloc[0]
                customer_fi_doc = int(dpc["DocumentNo"].iloc[0])
                customer_document_typ = dpc["Typ"].iloc[0]
                customer_posting_date = dpc["Pstng Date"].iloc[0].date().strftime("%Y-%m-%d")

                # --------------------------------------------------
                # INVOICE DETAILS
                # --------------------------------------------------
                invoice_number = dpc["Reference"].iloc[0]
                invoice_date = dpc["Doc..Date"].iloc[0].date().strftime("%Y-%m-%d")
                invoice_amount_rv = float(dpc["      Amount in DC"].iloc[0])

                remarks = "Customer DP Record which is not present in the Vendor sheet"

                not_matched_df.loc[len(not_matched_df),['Customer code',
                                                        'Customer FI Doc', 
                                                        'Customer/Vendor Document Type',
                                                        'Customer posting date',
                                                        'Invoice number', 
                                                        'Invoice date',
                                                        'invoice amount-rv',
                                                        'Remarks']] = [customer_code,
                                                                    customer_fi_doc,
                                                                    customer_document_typ,
                                                                    customer_posting_date,
                                                                    invoice_number,
                                                                    invoice_date,
                                                                    invoice_amount_rv,
                                                                    remarks]



            # vendor dp type records which are not present in the customer sheet
            len_of_vendor_dp_not_present = len(vendor_dp_records_which_are_not_in_customer)
            for cc in range(1,len_of_vendor_dp_not_present+1):
                dpv = vendor_dp_records_which_are_not_in_customer.iloc[cc-1:cc, :]

                # --------------------------------------------------
                # VENDOR DETAILS
                # --------------------------------------------------
                vendor_code = int(dpv["Vendor"].iloc[0])
                vendor_fi_doc = int(dpv["DocumentNo"].iloc[0])
                vendor_document_typ = dpv["Ty"].iloc[0]
                vendor_posting_date = dpv["Posting Date"].iloc[0].date().strftime('%Y-%m-%d')


                # --------------------------------------------------
                # INVOICE DETAILS
                # --------------------------------------------------
                invoice_number = dpv["Reference"].iloc[0]
                invoice_date = dpv["Doc--Date"].iloc[0].date().strftime("%Y-%m-%d")
                invoice_amount_re = abs(float(dpv["     Amount in LC"].iloc[0]))

                remarks = "Vendor DP Record which is not present in the Customer sheet"

                not_matched_df.loc[len(not_matched_df),['Vendor code', 
                                                        'Vendor FI Doc', 
                                                        'Vendor posting date',
                                                        'Customer/Vendor Document Type',
                                                        'Invoice number', 
                                                        'Invoice date',
                                                        'Posted amount-re',
                                                        'Remarks']] = [customer_code,
                                                                    customer_fi_doc,
                                                                    customer_document_typ,
                                                                    customer_posting_date,
                                                                    invoice_number,
                                                                    invoice_date,
                                                                    invoice_amount_rv,
                                                                    remarks]



            # =========================
            # WRITE TO EXCEL
            # =========================
            with pd.ExcelWriter(output_loc, engine="xlsxwriter") as writer:
                matched_df_final.to_excel(writer, sheet_name="matched", index=False)
                not_matched_df.to_excel(writer, sheet_name="not matched", index=False)

            # adding the customer records which are not present in the vendor sheet in the notmatched df
            len_of_customer_not_matched = len(customer_not_in_vendor)
            for bb in range(1,len_of_customer_not_matched+1):
                dpc = customer_not_in_vendor.iloc[bb-1:bb, :]


                # --------------------------------------------------
                # CUSTOMER DETAILS
                # --------------------------------------------------
                customer_code = dpc["Customer"].iloc[0]
                customer_fi_doc = int(dpc["DocumentNo_x"].iloc[0])
                customer_document_typ = dpc["Typ"].iloc[0]
                customer_posting_date = dpc["Pstng Date"].iloc[0].date().strftime("%Y-%m-%d")

                # --------------------------------------------------
                # INVOICE DETAILS
                # --------------------------------------------------
                invoice_number = dpc["Reference"].iloc[0]
                invoice_date = dpc["Doc..Date"].iloc[0].date().strftime("%Y-%m-%d")
                invoice_amount_rv = float(dpc["      Amount in DC"].iloc[0])

                remarks = "Customer Record which is not present in the Vendor sheet"

                not_matched_df.loc[len(not_matched_df),['Customer code',
                                                        'Customer FI Doc', 
                                                        'Customer/Vendor Document Type',
                                                        'Customer posting date',
                                                        'Invoice number', 
                                                        'Invoice date',
                                                        'invoice amount-rv',
                                                        'Remarks']] = [customer_code,
                                                                    customer_fi_doc,
                                                                    customer_document_typ,
                                                                    customer_posting_date,
                                                                    invoice_number,
                                                                    invoice_date,
                                                                    invoice_amount_rv,
                                                                    remarks]

            # =========================
            # WRITE TO EXCEL
            # =========================
            with pd.ExcelWriter(output_loc, engine="xlsxwriter") as writer:
                matched_df_final.to_excel(writer, sheet_name="matched", index=False)
                not_matched_df.to_excel(writer, sheet_name="not matched", index=False)


            # adding the vendor records which are not present in the customer sheet in the notmatched df
            len_of_vendor_not_matched = len(vendor_not_in_customer)
            for cc in range(1,len_of_vendor_not_matched+1):
                dpv = vendor_not_in_customer.iloc[cc-1:cc, :]

                # --------------------------------------------------
                # VENDOR DETAILS
                # --------------------------------------------------
                vendor_code = int(dpv["Vendor"].iloc[0])
                vendor_fi_doc = int(dpv["DocumentNo_x"].iloc[0])
                vendor_document_typ = dpv["Ty"].iloc[0]
                vendor_posting_date = dpv["Posting Date"].iloc[0].date().strftime('%Y-%m-%d')


                # --------------------------------------------------
                # INVOICE DETAILS
                # --------------------------------------------------
                invoice_number = dpv["Reference"].iloc[0]
                invoice_date = dpv["Doc--Date"].iloc[0].date().strftime("%Y-%m-%d")
                invoice_amount_re = abs(float(dpv["     Amount in LC"].iloc[0]))

                remarks = "Vendor Record which is not present in the Customer sheet"

                not_matched_df.loc[len(not_matched_df),['Vendor code', 
                                                        'Vendor FI Doc', 
                                                        'Vendor posting date',
                                                        'Customer/Vendor Document Type',
                                                        'Invoice number', 
                                                        'Invoice date',
                                                        'Posted amount-re',
                                                        'Remarks']] = [customer_code,
                                                                    customer_fi_doc,
                                                                    customer_document_typ,
                                                                    customer_posting_date,
                                                                    invoice_number,
                                                                    invoice_date,
                                                                    invoice_amount_rv,
                                                                    remarks]
            # =========================
            # WRITE TO EXCEL
            # =========================
            with pd.ExcelWriter(output_loc, engine="xlsxwriter") as writer:
                matched_df_final.to_excel(writer, sheet_name="matched", index=False)
                not_matched_df.to_excel(writer, sheet_name="not matched", index=False)


            vendor = vendor_ab_sa.copy()
            customer = customer_ab_sa.copy()

            ########################################################################################################################################################
            ########################################################################################################################################################
            ########################################################################################################################################################
            ########################################################################################################################################################
            ########################################################################################################################################################
            # customer AB doc type extraction
            customer_ab_records = customer[customer["Typ"].str.contains("AB",regex = False,na = False)]

            doc_stats = (
                customer_ab_records
                .groupby("DocumentNo", as_index=False)
                .agg(
                    record_count=("      Amount in DC", "count"),
                    amount_sum=("      Amount in DC", "sum")
                )
            )

            #################################################################
            #---------------------------------------------------------
            # which document count >=2 and the sum is = 0
            # so we have to store these records in the matched sheet
            #---------------------------------------------------------
            valid_docs = doc_stats[
                (doc_stats["record_count"] >= 2) &
                (doc_stats["amount_sum"] == 0)
            ]

            #################################################################
            #-------------------------------------------------------------
            # which document count is >=2 and the sum is not = 0
            # so we have to store these records in the not matched sheet
            #-------------------------------------------------------------
            invalid_docs = doc_stats[
                (doc_stats["record_count"] >= 2) &
                (doc_stats["amount_sum"] != 0)
            ]
            ########################################################################################################################################################


            zero_amount_docs = valid_docs
            AB_valid_df = customer_ab_records[
                customer_ab_records["DocumentNo"].isin(zero_amount_docs["DocumentNo"])
            ]

            for pp in range(1,len(AB_valid_df)+1):
                # Customer code
                ab_customer_code = AB_valid_df["Customer"].iloc[pp-1:pp]

                # Customer FI Doc
                ab_customer_fi_doc = AB_valid_df["DocumentNo"].iloc[pp-1:pp]

                # Customer Document Type
                ab_customer_document_type = AB_valid_df["Typ"].iloc[pp-1:pp]

                # Customer posting date
                ab_customer_posting_date = str(AB_valid_df["Pstng Date"].iloc[pp-1:pp].dt.date.iloc[0])
                # ab_customer_posting_date = ab_customer_posting_date.strftime("%Y-%m-%d")

                # invoice number
                ab_invoice_number = AB_valid_df["Reference"].iloc[pp-1:pp]

                # invoice date
                ab_invoice_date = str(AB_valid_df["Doc..Date"].iloc[pp-1:pp].dt.date.iloc[0])

                # invoice amount-rv
                ab_invoice_amount_rv = AB_valid_df["      Amount in DC"].iloc[pp-1:pp]

                # remarks_of_ab
                remarks_customer_valid_of_ab = "Matched with Doc Type AB, and sum = 0"


            ########################################################################################################################################################
                matched_df_final.loc[len(matched_df_final), ["Customer code",
                                                            "Customer FI Doc",
                                                            "Customer/Vendor Document Type",
                                                            "Customer posting date",
                                                            "Invoice number",
                                                            "Invoice date",
                                                            "invoice amount-rv",
                                                            "Remarks"]] = ab_customer_code.iloc[0],ab_customer_fi_doc.iloc[0],ab_customer_document_type.iloc[0],ab_customer_posting_date,ab_invoice_number.iloc[0],ab_invoice_date,ab_invoice_amount_rv.iloc[0],remarks_customer_valid_of_ab



            ########################################################################################################################################################
            ########################################################################################################################################################
            ########################################################################################################################################################
            ########################################################################################################################################################
            ########################################################################################################################################################
            # vendor AB doc type extraction
            vendor_ab_records = vendor[vendor["Ty"].str.contains("AB",regex = False,na = False)]

            doc_stats1 = (
                vendor_ab_records
                .groupby("DocumentNo", as_index=False)
                .agg(
                    record_count=("     Amount in LC", "count"),
                    amount_sum=("     Amount in LC", "sum")
                )
            )


            #################################################################
            #---------------------------------------------------------
            # which document count >=2 and the sum is = 0
            # so we have to store these records in the matched sheet
            #---------------------------------------------------------
            valid_vendor_docs = doc_stats1[
                (doc_stats1["record_count"] >= 2) &
                (doc_stats1["amount_sum"] == 0)
            ]

            #################################################################
            #-------------------------------------------------------------
            # which document count is >=2 and the sum is not = 0
            # so we have to store these records in the not matched sheet
            #-------------------------------------------------------------
            invalid_vendor_docs = doc_stats1[
                (doc_stats1["record_count"] >= 2) &
                (doc_stats1["amount_sum"] != 0)
            ]
            ########################################################################################################################################################


            zero_amount_docs1 = valid_vendor_docs
            AB_vendor_valid_df = vendor_ab_records[
                vendor_ab_records["DocumentNo"].isin(zero_amount_docs1["DocumentNo"])
            ]

            for ppp in range(1,len(AB_vendor_valid_df)+1):
                # Vendor code
                ab_vendor_code = AB_vendor_valid_df["Vendor"].iloc[ppp-1:ppp]

                # Vendor FI Doc
                ab_vendor_fi_doc = AB_vendor_valid_df["DocumentNo"].iloc[ppp-1:ppp]

                # Vendor Document Type
                ab_vendor_document_type = AB_vendor_valid_df["Ty"].iloc[ppp-1:ppp]

                # Vendor posting date
                ab_vendor_posting_date = str(AB_vendor_valid_df["Posting Date"].iloc[ppp-1:ppp].dt.date.iloc[0])

                # invoice number
                ab_vendor_invoice_number = AB_vendor_valid_df["Reference"].iloc[ppp-1:ppp]

                # invoice date
                ab_vendor_invoice_date = str(AB_vendor_valid_df["Doc--Date"].iloc[ppp-1:ppp].dt.date.iloc[0])

                # invoice amount-rv
                ab_vendor_invoice_amount_re = AB_vendor_valid_df["     Amount in LC"].iloc[ppp-1:ppp]

                # remarks_of_ab
                remarks_vendor_valid_of_ab = "Matched with Doc Type AB, and sum = 0"

            ########################################################################################################################################################
                matched_df_final.loc[len(matched_df_final), ["Vendor code",
                                                "Vendor FI Doc",
                                                "Vendor posting date",
                                                "Customer/Vendor Document Type",
                                                "Invoice number",
                                                "Invoice date",
                                                "posted amount-re",
                                                "Remarks"]] = ab_vendor_code.iloc[0],ab_vendor_fi_doc.iloc[0],ab_vendor_document_type.iloc[0],ab_vendor_posting_date,ab_vendor_invoice_number.iloc[0],ab_vendor_invoice_date,ab_vendor_invoice_amount_re.iloc[0],remarks_vendor_valid_of_ab

















            ########################################################################################################################################################
            ########################################################################################################################################################
            ########################################################################################################################################################
            ########################################################################################################################################################
            ########################################################################################################################################################
            # customer AB doc type extraction which document number wise not sum  = 0 
            customer_ab_records = customer[customer["Typ"].str.contains("AB",regex = False,na = False)]

            doc_stats11 = (
                customer_ab_records
                .groupby("DocumentNo", as_index=False)
                .agg(
                    record_count=("      Amount in DC", "count"),
                    amount_sum=("      Amount in DC", "sum")
                )
            )

            #################################################################
            #---------------------------------------------------------
            # which document count >=2 and the sum is = 0
            # so we have to store these records in the matched sheet
            #---------------------------------------------------------
            valid_docs = doc_stats11[
                (doc_stats11["record_count"] >= 2) &
                (doc_stats11["amount_sum"] == 0)
            ]

            #################################################################
            #-------------------------------------------------------------
            # which document count is >=2 and the sum is not = 0
            # so we have to store these records in the not matched sheet
            #-------------------------------------------------------------
            invalid_docs = doc_stats11[
                (doc_stats11["record_count"] >= 2) &
                (doc_stats11["amount_sum"] != 0)
            ]
            ########################################################################################################################################################


            zero_amount_docs11 = invalid_docs
            AB_invalid_df = customer_ab_records[
                customer_ab_records["DocumentNo"].isin(zero_amount_docs11["DocumentNo"])
            ]

            for pp1 in range(1,len(AB_invalid_df)+1):
                # Customer code
                ab_customer_code1 = AB_invalid_df["Customer"].iloc[pp1-1:pp1]

                # Customer FI Doc
                ab_customer_fi_doc1 = AB_invalid_df["DocumentNo"].iloc[pp1-1:pp1]

                # Customer Document Type
                ab_customer_document_type1 = AB_invalid_df["Typ"].iloc[pp1-1:pp1]

                # Customer posting date
                ab_customer_posting_date1 = str(AB_invalid_df["Pstng Date"].iloc[pp1-1:pp1].dt.date.iloc[0])

                # invoice number
                ab_invoice_number1 = AB_invalid_df["Reference"].iloc[pp1-1:pp1]

                # invoice date
                ab_invoice_date1 = str(AB_invalid_df["Doc..Date"].iloc[pp1-1:pp1].dt.date.iloc[0])

                # invoice amount-rv
                ab_invoice_amount_rv1 = AB_invalid_df["      Amount in DC"].iloc[pp1-1:pp1]

                # remarks_of_ab
                remarks_customer_invalid_of_ab = "Doc Type AB, and sum not = 0"


            ########################################################################################################################################################
                not_matched_df.loc[len(not_matched_df), ["Customer code",
                                                        "Customer FI Doc",
                                                        "Customer/Vendor Document Type",
                                                        "Customer posting date",
                                                        "Invoice number",
                                                        "Invoice date",
                                                        "invoice amount-rv",
                                                        "Remarks"]] = [ab_customer_code1.iloc[0],
                                                                        ab_customer_fi_doc1.iloc[0],
                                                                        ab_customer_document_type1.iloc[0],
                                                                        ab_customer_posting_date1,
                                                                        ab_invoice_number1.iloc[0],
                                                                        ab_invoice_date1,
                                                                        ab_invoice_amount_rv1.iloc[0],
                                                                        remarks_customer_invalid_of_ab
                                                                    ]


            ########################################################################################################################################################
            ########################################################################################################################################################
            ########################################################################################################################################################
            ########################################################################################################################################################
            ########################################################################################################################################################
            # vendor AB doc type extraction which document number wise sum != 0 
            vendor_ab_records = vendor[vendor["Ty"].str.contains("AB",regex = False,na = False)]

            doc_stats111 = (
                vendor_ab_records
                .groupby("DocumentNo", as_index=False)
                .agg(
                    record_count=("     Amount in LC", "count"),
                    amount_sum=("     Amount in LC", "sum")
                )
            )


            #################################################################
            #---------------------------------------------------------
            # which document count >=2 and the sum is = 0
            # so we have to store these records in the matched sheet
            #---------------------------------------------------------
            valid_vendor_docs = doc_stats111[
                (doc_stats111["record_count"] >= 2) &
                (doc_stats111["amount_sum"] == 0)
            ]

            #################################################################
            #-------------------------------------------------------------
            # which document count is >=2 and the sum is not = 0
            # so we have to store these records in the not matched sheet
            #-------------------------------------------------------------
            invalid_vendor_docs = doc_stats111[
                (doc_stats111["record_count"] >= 2) &
                (doc_stats111["amount_sum"] != 0)
            ]
            ########################################################################################################################################################


            zero_amount_docs111 = invalid_vendor_docs
            AB_vendor_valid_df1 = vendor_ab_records[
                vendor_ab_records["DocumentNo"].isin(zero_amount_docs111["DocumentNo"])
            ]

            for ppp in range(1,len(AB_vendor_valid_df1)+1):
                # Vendor code
                ab_vendor_code1 = AB_vendor_valid_df1["Vendor"].iloc[ppp-1:ppp]

                # Vendor FI Doc
                ab_vendor_fi_doc1 = AB_vendor_valid_df1["DocumentNo"].iloc[ppp-1:ppp]

                # Vendor Document Type
                ab_vendor_document_type1 = AB_vendor_valid_df1["Ty"].iloc[ppp-1:ppp]

                # Vendor posting date
                ab_vendor_posting_date1 = str(AB_vendor_valid_df1["Posting Date"].iloc[ppp-1:ppp].dt.date.iloc[0])

                # invoice number
                ab_vendor_invoice_number1 = AB_vendor_valid_df1["Reference"].iloc[ppp-1:ppp]

                # invoice date
                ab_vendor_invoice_date1 = str(AB_vendor_valid_df1["Doc--Date"].iloc[ppp-1:ppp].dt.date.iloc[0])

                # invoice amount-rv
                ab_vendor_invoice_amount_re1 = AB_vendor_valid_df1["     Amount in LC"].iloc[ppp-1:ppp]

                # remarks_of_ab
                remarks_vendor_invalid_of_ab = "Type AB, and sum not = 0"
            ########################################################################################################################################################
                not_matched_df.loc[len(not_matched_df), ["Vendor code",
                                                        "Vendor FI Doc",
                                                        "Vendor posting date",
                                                        "Customer/Vendor Document Type",
                                                        "Invoice number",
                                                        "Invoice date",
                                                        "posted amount-re",
                                                        "Remarks"]] = [ab_vendor_code1.iloc[0],
                                                                        ab_vendor_fi_doc1.iloc[0],
                                                                        ab_vendor_document_type1.iloc[0],
                                                                        ab_vendor_posting_date1,
                                                                        ab_vendor_invoice_number1.iloc[0],
                                                                        ab_vendor_invoice_date1,
                                                                        ab_vendor_invoice_amount_re1.iloc[0],
                                                                        remarks_vendor_invalid_of_ab
                                                                    ]



            # write to excel
            with pd.ExcelWriter(
                output_loc,
                engine="xlsxwriter"
            ) as writer:
                matched_df_final.to_excel(writer, sheet_name="matched", index=False)
                not_matched_df.to_excel(writer, sheet_name="not matched", index=False)


            # Customer AB records which document have 1 records
            customer_ab_records_which_have_1_record = doc_stats[(doc_stats["record_count"] == 1)]
            customer_ab_records_which_have_1_record_row = customer[customer["DocumentNo"].isin(customer_ab_records_which_have_1_record["DocumentNo"])]


            # vendor AB records which document have 1 records
            vendor_ab_records_which_have_1_record = (doc_stats1[(doc_stats1["record_count"] == 1)])
            vendor_ab_records_which_have_1_record_row = vendor[vendor["DocumentNo"].isin(vendor_ab_records_which_have_1_record["DocumentNo"])]



            # adding the customer ab records which document number wise count is 1
            len_of_customer_ab_singl_record = len(customer_ab_records_which_have_1_record_row)
            for cc in range(1,len_of_customer_ab_singl_record+1):
                dpc = customer_ab_records_which_have_1_record_row.iloc[cc-1:cc, :]


                # --------------------------------------------------
                # CUSTOMER DETAILS
                # --------------------------------------------------
                customer_code = dpc["Customer"].iloc[0]
                customer_fi_doc = int(dpc["DocumentNo"].iloc[0])
                customer_document_typ = dpc["Typ"].iloc[0]
                customer_posting_date = dpc["Pstng Date"].iloc[0].date().strftime("%Y-%m-%d")

                # --------------------------------------------------
                # INVOICE DETAILS
                # --------------------------------------------------
                invoice_number = dpc["Reference"].iloc[0]
                invoice_date = dpc["Doc..Date"].iloc[0].date().strftime("%Y-%m-%d")
                invoice_amount_rv = float(dpc["      Amount in DC"].iloc[0])

                remarks = "Customer AB Record, which has single data point"

                not_matched_df.loc[len(not_matched_df),['Customer code',
                                                        'Customer FI Doc', 
                                                        'Customer/Vendor Document Type',
                                                        'Customer posting date',
                                                        'Invoice number', 
                                                        'Invoice date',
                                                        'invoice amount-rv',
                                                        'Remarks']] = [customer_code,
                                                                    customer_fi_doc,
                                                                    customer_document_typ,
                                                                    customer_posting_date,
                                                                    invoice_number,
                                                                    invoice_date,
                                                                    invoice_amount_rv,
                                                                    remarks]

            # =========================
            # WRITE TO EXCEL
            # =========================
            with pd.ExcelWriter(output_loc, engine="xlsxwriter") as writer:
                matched_df_final.to_excel(writer, sheet_name="matched", index=False)
                not_matched_df.to_excel(writer, sheet_name="not matched", index=False)

            # adding the vendor ab records which document number wise count is 1
            len_of_vendor_ab_singl_record = len(vendor_ab_records_which_have_1_record_row)
            for dd in range(1,len_of_vendor_ab_singl_record+1):
                dpv = vendor_ab_records_which_have_1_record_row.iloc[dd-1:dd, :]

                # --------------------------------------------------
                # VENDOR DETAILS
                # --------------------------------------------------
                vendor_code = int(dpv["Vendor"].iloc[0])
                vendor_fi_doc = int(dpv["DocumentNo"].iloc[0])
                vendor_document_typ = dpv["Ty"].iloc[0]
                vendor_posting_date = dpv["Posting Date"].iloc[0].date().strftime('%Y-%m-%d')


                # --------------------------------------------------
                # INVOICE DETAILS
                # --------------------------------------------------
                invoice_number = dpv["Reference"].iloc[0]
                invoice_date = dpv["Doc--Date"].iloc[0].date().strftime("%Y-%m-%d")
                invoice_amount_re = abs(float(dpv["     Amount in LC"].iloc[0]))

                remarks = "Vendor Record which is not present in the Customer sheet"

                not_matched_df.loc[len(not_matched_df),['Vendor code', 
                                                        'Vendor FI Doc', 
                                                        'Vendor posting date',
                                                        'Customer/Vendor Document Type',
                                                        'Invoice number', 
                                                        'Invoice date',
                                                        'Posted amount-re',
                                                        'Remarks']] = [customer_code,
                                                                    customer_fi_doc,
                                                                    customer_document_typ,
                                                                    customer_posting_date,
                                                                    invoice_number,
                                                                    invoice_date,
                                                                    invoice_amount_rv,
                                                                    remarks]
            # =========================
            # WRITE TO EXCEL
            # =========================
            with pd.ExcelWriter(output_loc, engine="xlsxwriter") as writer:
                matched_df_final.to_excel(writer, sheet_name="matched", index=False)
                not_matched_df.to_excel(writer, sheet_name="not matched", index=False)

            vendor1 = vendor_ab_sa.copy()
            customer1 = customer_ab_sa.copy()


            #####################################################################################################################################################
            # new excel sheets
            sa_vendor_df = pd.DataFrame(columns=[
                "Vendor code",
                "Vendor FI Doc",
                "Vendor posting date",
                "Vendor Document Type",
                "Invoice number",
                "Invoice date",
                "Posted amount-re",
                "Status"
            ])


            #####################################################################################################################################################
            sa_customer_df = pd.DataFrame(columns=[
                "Customer code",
                "Customer FI Doc",
                "Customer Posting Date",
                "Customer Document Type",
                "Invoice number",
                "Invoice date",
                "Posted amount-rv",
                "Status"
            ])



            #####################################################################################################################################################
            vendor1 = vendor_ab_sa.copy()
            customer1 = customer_ab_sa.copy()

            ###################################################################################################################################################
            # filter the data points which have document type as SA
            sa_vendor_records = vendor1[vendor1["Ty"].str.contains("SA",regex = False, na=False)]

            # filter the data points which have document type as SA
            sa_customer_records = customer1[customer1["Typ"].str.contains("SA",regex = False, na = False)]

            # only taking the records of customer excel which documents number is matched with the vendor excel
            vendor_sa_document_number_matched_with_customer_document_number = sa_vendor_records[
                                                                                            sa_vendor_records["DocumentNo"].isin(
                                                                                                    sa_customer_records["DocumentNo"])]

            ####################################################################################################################################################
            ####################################################################################################################################################
            ####################################################################################################################################################
            #############################################################
            # --------------------------------------------------
            # 1. Group vendor
            # --------------------------------------------------
            dpv_grp = (
                sa_vendor_records
                .groupby("DocumentNo", as_index=False)
                .agg(
                    df1_sum=("     Amount in LC", "sum"),
                    df1_count=("     Amount in LC", "count")
                )
            )

            #############################################################
            # --------------------------------------------------
            # 2. Group df2
            # --------------------------------------------------
            dpc_grp = (
                sa_customer_records
                .groupby("DocumentNo", as_index=False)
                .agg(
                    df2_sum=("      Amount in DC", "sum"),
                    df2_count=("      Amount in DC", "count")
                )
            )

            #############################################################
            # --------------------------------------------------
            # 3. Merge
            # --------------------------------------------------
            merged = dpv_grp.merge(
                dpc_grp,
                on="DocumentNo",
                how="outer"
            )

            # Replace NaNs
            merged[["df1_sum", "df2_sum", "df1_count", "df2_count"]] = (
                merged[["df1_sum", "df2_sum", "df1_count", "df2_count"]]
                .fillna(0)
            )


            #############################################################
            # --------------------------------------------------
            # 4. Apply your business logic
            # --------------------------------------------------
            merged["difference"] = abs(merged["df1_sum"] + (merged["df2_sum"]))

            merged["difference"] = merged["difference"].astype(float)
            ###################################################################################################################################################
            ###################################################################################################################################################
            ###################################################################################################################################################

            # sa_matched_docs = merged.loc[merged["difference"] == 0, "DocumentNo"]
            # sa_notmatched_docs = merged.loc[merged["difference"] != 0, "DocumentNo"]
            sa_matched_docs = merged.loc[
                (merged["difference"] >= 0) & (merged["difference"] <= 10),
                "DocumentNo"
            ]

            sa_notmatched_docs = merged.loc[
                (merged["difference"] < 0) | (merged["difference"] > 10), 
                "DocumentNo"
            ]


            ##################################################################
            sa_vendor_matched_rows = sa_vendor_records[
                sa_vendor_records["DocumentNo"].isin(sa_matched_docs)
            ]

            sa_vendor_notmatched_rows = sa_vendor_records[
                sa_vendor_records["DocumentNo"].isin(sa_notmatched_docs)
            ]


            ##################################################################
            sa_customer_matched_rows = sa_customer_records[
                sa_customer_records["DocumentNo"].isin(sa_matched_docs)
            ]

            sa_customer_notmatched_rows = sa_customer_records[
                sa_customer_records["DocumentNo"].isin(sa_notmatched_docs)
            ]

            ##################################################################
            ###################################################################################################################################################
            #-------------------------
            # sa_vendor_matched_df
            #-------------------------
            for q in range(len(sa_vendor_matched_rows)):

                sa_vendor_code = sa_vendor_matched_rows["Vendor"].iloc[q]
                sa_vendor_fi_doc = sa_vendor_matched_rows["DocumentNo"].iloc[q]
                sa_vendor_document_type = sa_vendor_matched_rows["Ty"].iloc[q]

                sa_vendor_posting_date = sa_vendor_matched_rows["Posting Date"].iloc[q]
                if pd.notna(sa_vendor_posting_date):
                    sa_vendor_posting_date = pd.to_datetime(sa_vendor_posting_date).strftime("%Y-%m-%d")

                sa_vendor_invoice_number = sa_vendor_matched_rows["Reference"].iloc[q]

                sa_vendor_invoice_date = sa_vendor_matched_rows["Doc--Date"].iloc[q]
                if pd.notna(sa_vendor_invoice_date):
                    sa_vendor_invoice_date = pd.to_datetime(sa_vendor_invoice_date).strftime("%Y-%m-%d")

                sa_vendor_invoice_amount_re = sa_vendor_matched_rows["     Amount in LC"].iloc[q]

                sa_vendor_matched_status = "Matched"

            #############################################################################################################
                sa_vendor_df.loc[len(sa_vendor_df)]= [
                    sa_vendor_code,
                    sa_vendor_fi_doc,
                    sa_vendor_posting_date,
                    sa_vendor_document_type,
                    sa_vendor_invoice_number,
                    sa_vendor_invoice_date,
                    sa_vendor_invoice_amount_re,
                    sa_vendor_matched_status
                ]

            ##################################################################################################################################################
            ##################################################################################################################################################
            #-------------------------
            # sa_vendor_notmatched_df
            #-------------------------
            for q in range(len(sa_vendor_notmatched_rows)):

                sa_vendor_code = sa_vendor_notmatched_rows["Vendor"].iloc[q]
                sa_vendor_fi_doc = sa_vendor_notmatched_rows["DocumentNo"].iloc[q]
                sa_vendor_document_type = sa_vendor_notmatched_rows["Ty"].iloc[q]

                sa_vendor_posting_date = sa_vendor_notmatched_rows["Posting Date"].iloc[q]
                if pd.notna(sa_vendor_posting_date):
                    sa_vendor_posting_date = pd.to_datetime(sa_vendor_posting_date).strftime("%Y-%m-%d")

                sa_vendor_invoice_number = sa_vendor_notmatched_rows["Reference"].iloc[q]

                sa_vendor_invoice_date = sa_vendor_notmatched_rows["Doc--Date"].iloc[q]
                if pd.notna(sa_vendor_invoice_date):
                    sa_vendor_invoice_date = pd.to_datetime(sa_vendor_invoice_date).strftime("%Y-%m-%d")

                sa_vendor_invoice_amount_re = sa_vendor_notmatched_rows["     Amount in LC"].iloc[q]

                sa_vendor_notmatched_status = "Not Matched"

            #############################################################################################################
                sa_vendor_df.loc[len(sa_vendor_df)]= [
                    sa_vendor_code,
                    sa_vendor_fi_doc,
                    sa_vendor_posting_date,
                    sa_vendor_document_type,
                    sa_vendor_invoice_number,
                    sa_vendor_invoice_date,
                    sa_vendor_invoice_amount_re,
                    sa_vendor_notmatched_status
                ]


            ###################################################################################################################################################
            ###################################################################################################################################################
            ###################################################################################################################################################
            ###################################################################################################################################################
            ###################################################################################################################################################
            ###################################################################################################################################################
            ###################################################################################################################################################
            ###################################################################################################################################################
            #-------------------------
            # sa_customer_matched_df
            #-------------------------
            for q in range(len(sa_customer_matched_rows)):
                #########################################################################################################
                # col1 (Customer code)
                sa_customer_code = sa_customer_matched_rows["Customer"].iloc[q]

                #########################################################################################################
                # col2 (Customer FI Doc)
                sa_customer_fi_doc = sa_customer_matched_rows["DocumentNo"].iloc[q]

                #########################################################################################################
                # col3 (Customer Posting Date)
                sa_customer_posting_date = sa_customer_matched_rows["Pstng Date"].iloc[q]
                if pd.notna(sa_customer_posting_date):
                    sa_customer_posting_date = pd.to_datetime(sa_customer_posting_date).strftime("%Y-%m-%d")

                #########################################################################################################
                # col 4 (Customer Document Type)
                sa_customer_document_type = sa_customer_matched_rows["Typ"].iloc[q]

                #########################################################################################################
                # col 5 (Invoice number)
                sa_customer_invoice_number = sa_customer_matched_rows["Reference"].iloc[q]

                #########################################################################################################
                # col 6 (Invoice date) 
                sa_customer_invoice_date = sa_customer_matched_rows["Doc..Date"].iloc[q]
                if pd.notna(sa_customer_invoice_date):
                    sa_customer_invoice_date = pd.to_datetime(sa_customer_invoice_date).strftime("%Y-%m-%d")

                #########################################################################################################
                # col 7 (Posted amount-rv)
                sa_customer_invoice_amount_rv = sa_customer_matched_rows["      Amount in DC"].iloc[q]

                #########################################################################################################
                # col 8 (Status)
                sa_customer_matched_status = "Matched"

                ##################################################################################################################
                sa_customer_df.loc[len(sa_customer_df)]= [
                    sa_customer_code,
                    sa_customer_fi_doc,
                    sa_customer_posting_date,
                    sa_customer_document_type,
                    sa_customer_invoice_number,
                    sa_customer_invoice_date,
                    sa_customer_invoice_amount_rv,
                    sa_customer_matched_status
                ]

            ###################################################################################################################################################
            ###################################################################################################################################################
            ###################################################################################################################################################
            ###################################################################################################################################################
            ###################################################################################################################################################
            ###################################################################################################################################################
            #-------------------------
            # sa_customer_notmatched_df
            #-------------------------
            for q in range(len(sa_customer_notmatched_rows)):
                #########################################################################################################
                # col1 (Customer code)
                sa_customer_code = sa_customer_notmatched_rows["Customer"].iloc[q]

                #########################################################################################################
                # col2 (Customer FI Doc)
                sa_customer_fi_doc = sa_customer_notmatched_rows["DocumentNo"].iloc[q]

                #########################################################################################################
                # col3 (Customer Posting Date)
                sa_customer_posting_date = sa_customer_notmatched_rows["Pstng Date"].iloc[q]
                if pd.notna(sa_customer_posting_date):
                    sa_customer_posting_date = pd.to_datetime(sa_customer_posting_date).strftime("%Y-%m-%d")

                #########################################################################################################
                # col 4 (Customer Document Type)
                sa_customer_document_type = sa_customer_notmatched_rows["Typ"].iloc[q]

                #########################################################################################################
                # col 5 (Invoice number)
                sa_customer_invoice_number = sa_customer_notmatched_rows["Reference"].iloc[q]

                #########################################################################################################
                # col 6 (Invoice date) 
                sa_customer_invoice_date = sa_customer_notmatched_rows["Doc..Date"].iloc[q]
                if pd.notna(sa_customer_invoice_date):
                    sa_customer_invoice_date = pd.to_datetime(sa_customer_invoice_date).strftime("%Y-%m-%d")

                #########################################################################################################
                # col 7 (Posted amount-rv)
                sa_customer_invoice_amount_rv = sa_customer_notmatched_rows["      Amount in DC"].iloc[q]

                #########################################################################################################
                # col 8 (Status)
                sa_customer_notmatched_status = "Not Matched"
                ##################################################################################################################
                sa_customer_df.loc[len(sa_customer_df)]= [
                    sa_customer_code,
                    sa_customer_fi_doc,
                    sa_customer_posting_date,
                    sa_customer_document_type,
                    sa_customer_invoice_number,
                    sa_customer_invoice_date,
                    sa_customer_invoice_amount_rv,
                    sa_customer_notmatched_status
                ]
            ####################################################################################################################################################
            ####################################################################################################################################################
            ####################################################################################################################################################
            #-------------------   
            # write to excel
            #-------------------

            with pd.ExcelWriter(
                output_loc,
                engine="xlsxwriter"
            ) as writer:
                matched_df_final.to_excel(writer, sheet_name="Matched", index=False)
                not_matched_df.to_excel(writer, sheet_name="Not matched", index=False)
                sa_vendor_df.to_excel(writer, sheet_name="Vendor SA", index=False)
                sa_customer_df.to_excel(writer, sheet_name="Customer SA", index=False)






            ###################################################################################################
            ###################################################################################################
            time.sleep(3)





            output_loc.seek(0)

        st.success("Processing completed!")

        # =========================
        # DOWNLOAD OUTPUT
        # =========================
        st.download_button(
            label="Download Output File",
            data=output_loc,
            file_name="output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# =========================
# FOOTER
# =========================
st.markdown("""
<div class="footer">
    © 2026 Jindal Stainless Steel. All rights reserved.
</div>
""", unsafe_allow_html=True)











