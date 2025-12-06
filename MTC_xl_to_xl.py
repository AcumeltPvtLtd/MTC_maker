import openpyxl
import os

# --- Configuration: File Paths ---
# Using the specific paths from your environment
SOURCE_PATH = "/home/johnny/Downloads/Book46.xlsx"
DEST_PATH = "/home/johnny/MTCAUTO/MTCAUTO/MTC_R3(HUB-821108)ENTRY_SHEET-2025.xlsx"

# --- Configuration: Mapping ---
# Format: [Source_Cell (Book46), Destination_Cell (MTC_R3)]
# Logic: Take value FROM Source, Write TO Destination
CELL_MAPPINGS = [
    ['C2',  'E14'],  # Heat Code?
    ['B10', 'E16'],  # Element 1
    ['C10', 'E17'],  # Element 2
    ['D10', 'E18'],  # Element 3
    ['F10', 'E19'],  # Element 4 (Skipped E in source)
    ['P10', 'E20'],  # Element 5 (Updated: P10 instead of E10)
    ['G10', 'E21'],  # Element 6
    ['L10', 'E22'],  # Element 7
    ['S10', 'E23']   # Element 8
]

def transfer_specific_data():
    print("--- Starting Exact Cell Transfer ---")

    # 1. Check if files exist
    if not os.path.exists(SOURCE_PATH):
        print(f"❌ Error: Source file not found at: {SOURCE_PATH}")
        return
    if not os.path.exists(DEST_PATH):
        print(f"❌ Error: Destination file not found at: {DEST_PATH}")
        return

    try:
        # 2. Load Source Workbook
        # data_only=True ensures we get the calculated value, not the formula string
        print(f"📂 Loading Source: {os.path.basename(SOURCE_PATH)}...")
        wb_source = openpyxl.load_workbook(SOURCE_PATH, data_only=True)
        ws_source = wb_source.active 

        # 3. Load Destination Workbook
        print(f"📂 Loading Destination: {os.path.basename(DEST_PATH)}...")
        wb_dest = openpyxl.load_workbook(DEST_PATH)
        ws_dest = wb_dest.active 

        print("\n🔄 Transferring Values...")

        # 4. Iterate through the mapping list and transfer data
        transfer_count = 0
        for pair in CELL_MAPPINGS:
            source_cell = pair[0]
            dest_cell = pair[1]

            # Get value
            value = ws_source[source_cell].value

            # Write value
            ws_dest[dest_cell].value = value
            
            print(f"   ✅ [Map {transfer_count+1}] Copied '{value}' from {source_cell} (Source) to {dest_cell} (Dest)")
            transfer_count += 1

        # 5. Save the Destination File
        # We overwrite the destination file (DEST_PATH) as implied by your request,
        # or you can change this to a new name like "MTC_Updated.xlsx" to be safe.
        print(f"\n💾 Saving file to: {DEST_PATH}")
        wb_dest.save(DEST_PATH)
        print("🎉 Transfer Complete!")

    except Exception as e:
        print(f"❌ An unexpected error occurred: {e}")

if __name__ == "__main__":
    transfer_specific_data()
