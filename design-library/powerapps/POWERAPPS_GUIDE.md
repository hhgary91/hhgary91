# Design Library Search in Power Apps: setup guide

This guide builds the search tool as a Power Apps canvas app. The app reads `Design_Library_Register.xlsx` directly from your SharePoint library, and you put it on a SharePoint page. You don't need IT approval.

**Files in this folder**

| File | What it is |
|---|---|
| `01_Load_Data.fx.txt` | The formula that loads all seven registers into one searchable list. It's already inside the search screen; this copy is for reference. |
| `02_scrSearch.pa.yaml` | The search screen: search box, three filters and a results list. You paste this into Power Apps. |
| `03_scrDetail.pa.yaml` | The detail screen: design summary, checks, file buttons and rebar schedule. You paste this into Power Apps. |

> These screens were written outside Power Apps and have not been tested in your company's setup yet. If a paste is refused, see **Troubleshooting** at the end. Every formula still works if you type it in by hand.

---

## Step 1. Put the register in SharePoint

1. Upload `Design_Library_Register.xlsx` to your SharePoint document library, for example into a folder named `00_Register`.
2. Open it in Excel and go to the **SG_Criteria** sheet. Set these two cells:
   - **Library root folder**: the synced path on your PC. Used by AutoCAD and Excel.
   - **Library web address (SharePoint)**: the browser address of the same library. To find it, open the library in your browser and copy the address bar. Remove anything from `/Forms/...` onwards. Example: `https://yourcompany.sharepoint.com/sites/Engineering/Design Library`
3. **Save** the file in Excel. Power Apps reads the values Excel saved, so the file must have been saved by Excel at least once.
4. Do not rename the tables (`tblDrain`, `tblRW`, …) or the column headings. The app uses those names.

## Step 2. Create a blank app

1. Go to **make.powerapps.com** and sign in with your work account.
2. Click **+ Create** → **Blank app** → **Blank canvas app**.
3. Name it `Design Library Search`, choose **Tablet** format, and click **Create**.

## Step 3. Connect the Excel tables

1. In the left bar, click the **Data** icon (the cylinder) → **Add data**.
2. Search for **Excel Online (Business)** and select it. Create a connection if asked.
3. Pick **SharePoint Site** → your site → the document library → `Design_Library_Register.xlsx`.
4. Tick these tables, then click **Connect**:
   `tblDrain, tblRW, tblEmbedded, tblManhole, tblERSSPit, tblLinkway, tblOther, tblRebar, tblProjects, tblRevisions`
5. Power Apps may add a hidden column called `__PowerAppsId__` to each table. That's normal and harmless.

## Step 4. Raise the row limit

**Settings** (gear icon) → **General** → **Data row limit**: set it to **2000**. Without this, Power Apps reads only the first 500 rows of each table.

## Step 5. Build the search screen

1. In the **Tree view**, rename `Screen1` to `scrSearch` (double-click the name).
2. Open `02_scrSearch.pa.yaml` in Notepad, select all and copy.
3. In Power Apps, right-click `scrSearch` in the Tree view and choose **Paste code**. (In some versions you select the screen and press **Ctrl+V**.)
4. Select `scrSearch`. In the property list (top left), choose **OnVisible** and enter:
   ```
   If(IsEmpty(colDesigns), Select(btnRefresh))
   ```

## Step 6. Build the detail screen

1. **+ New screen** → **Blank**, then rename it `scrDetail`.
2. Copy all of `03_scrDetail.pa.yaml` and paste it onto `scrDetail` the same way.

## Step 7. Add the full design record (about 5 minutes)

The detail screen shows the key information. To also show every field of a design, add one **display form** per register. Power Apps builds the field list for you.

For each row of the table below:
1. On `scrDetail`, choose **Insert** → **Display form**.
2. Rename it and set **DataSource**, **Item** and **Visible** as shown.
3. Set **X** = `600`, **Y** = `372`, **Width** = `Parent.Width - 624`, **Height** = `Parent.Height - 388`, **Columns** = `3`.
4. In the right panel, click **Edit fields** → **Add field**, tick all fields, then click **Add**. Remove any you don't want, such as `Selector Key`.

| Form name | DataSource | Item | Visible |
|---|---|---|---|
| frmDrain | `tblDrain` | `LookUp(tblDrain, 'Design ID' = gblSel.ID)` | `gblSel.Register = "Drain"` |
| frmRW | `tblRW` | `LookUp(tblRW, 'Design ID' = gblSel.ID)` | `gblSel.Register = "RC wall"` |
| frmEmbedded | `tblEmbedded` | `LookUp(tblEmbedded, 'Design ID' = gblSel.ID)` | `gblSel.Register = "Embedded wall"` |
| frmManhole | `tblManhole` | `LookUp(tblManhole, 'Design ID' = gblSel.ID)` | `gblSel.Register = "Manhole"` |
| frmERSSPit | `tblERSSPit` | `LookUp(tblERSSPit, 'Design ID' = gblSel.ID)` | `gblSel.Register = "ERSS pit"` |
| frmLinkway | `tblLinkway` | `LookUp(tblLinkway, 'Design ID' = gblSel.ID)` | `gblSel.Register = "Linkway"` |
| frmOther | `tblOther` | `LookUp(tblOther, 'Design ID' = gblSel.ID)` | `gblSel.Register = "Others"` |

## Step 8. Test

1. Press **F5** (Play).
2. Type `culvert` in the search box. The example box culverts should appear.
3. Click one. The detail screen should open, and **CAD detail** / **Drawing (.pdf)** should open the files from SharePoint once they're in the design folders.
4. Press **Esc** to stop.

## Step 9. Save, publish and share

1. **File** → **Save**, then **Publish**.
2. **Share** the app with your team, and make sure they can open `Design_Library_Register.xlsx` in SharePoint. The app uses each person's own access to the file. Read access is normally enough for searching; if someone gets a connection error, give them edit access.
3. The Excel Online (Business) and SharePoint connectors are standard connectors, normally covered by Microsoft 365 licences. If your company has special Power Apps policies, check with IT.

## Step 10. Put it on a SharePoint page

1. In make.powerapps.com, open the app's **Details** and copy the **Web link** (or the **App ID**).
2. Open your SharePoint page → **Edit** → **+** → **Power Apps** web part.
3. Paste the link or ID, set the height to about **800 px**, and **Republish** the page.

---

## Day-to-day use

- **Add or change designs in Excel** (desktop or Excel for the web), then **save**. The SG checks, IDs and completeness are recalculated by Excel when you save.
- In the app, click **Reload register** to see the changes.
- Avoid having two people editing the register at the same time. Use Excel for the web's co-authoring, or agree who edits.

## Troubleshooting

| Problem | Fix |
|---|---|
| **Paste code** is greyed out or the paste is refused | Make sure the screen is selected and empty. If Power Apps reports an unknown control version, delete the `@2.5.1`-style version numbers from the `Control:` lines in the file and paste again. As a last resort, insert the controls by hand (label, text input, three drop-downs, vertical gallery) using the names in the file, and copy each formula from the file into the matching property. |
| A column name is underlined red (e.g. `r.'Detail Scale (1:n)'`) | The Excel connector sometimes renames columns with special characters. Delete the name after `r.`, type `'` and pick the matching column from the list. |
| Description, compliance or completeness is blank | The file was changed without Excel saving it. Open it in Excel, save, then click **Reload register**. |
| Yellow "delegation" warning on `Filter(tblRebar, …)` | Safe to ignore while the rebar schedule has fewer than 2,000 rows. |
| File buttons open the wrong place | Check **Library web address** on SG_Criteria, and that each design's files sit in `<register folder>/<Design ID>/`, named `<Design ID>_CAD.dwg`, `_DWG.pdf`, `_STAAD.std`, `_CALC.xlsx` and `_CALC.pdf`. |
| App is slow to load | Normal for the first load. If the library grows to thousands of designs, move the registers into SharePoint Lists (ask Claude for the steps). |
