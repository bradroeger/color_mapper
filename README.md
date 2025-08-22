# 📘 Color Mapper – README

This tool looks at an **Excel file** and finds cells that are filled with certain colors.  
When it finds a match, it writes some text (you decide what) into the same row or column.  

You don’t need to know coding – just follow the steps.

---

## 🛠 What You Need
- **Windows computer**  
- **Your Excel file (.xlsx)**  
- **The Color Mapper program** (`color_mapper.exe`)  
- **A config file** (`config.json`) that tells the program:  
  - which colors to look for  
  - what text to write

---

## 📂 Folder Setup
Put everything in the same folder:
```
/ColorMapper
  ├─ color_mapper.exe     (the program)
  ├─ input.xlsx           (your Excel file)
  └─ config.json          (your settings file)
```

---

## ✍️ Editing the Config File
Open `config.json` in Notepad. It will look like this:

```json
{
  "FFFF0000": "This one is red",
  "FF00FF00": "This one is green",
  "FF0000FF": "This one is blue",
  "indexed:64": "No fill (white)",
  "theme:0": "Theme color 0"
}
```

- The **left side** (`FFFF0000`) is the **color code**.  
  - Example: `FFFF0000` = bright red  
  - Example: `FF00FF00` = green  
  - Example: `FF0000FF` = blue  
- The **right side** (`"This one is red"`) is the text that will be written in Excel.  

👉 You can add or remove as many colors as you like.  

---

## ▶️ How to Run It
1. Put your Excel file (`input.xlsx`) and `config.json` in the same folder as `color_mapper.exe`.  
2. Double-click **color_mapper.exe**.  
   - A black window will open.  
   - Type this and press Enter:
     ```
     color_mapper.exe input.xlsx config.json
     ```
3. The program will scan your file and create a new one:
   ```
   input_output.xlsx
   ```

---

## ⚙️ Options
You can change how the notes are added:

- **Row mode** (default if you don’t change anything):  
  Adds text in a new column on the right side of the row.

- **Column mode**:  
  Adds text **below** the last filled cell in the same column.  

To use column mode:
```
color_mapper.exe input.xlsx config.json --mode column
```

---

## 🔍 Debug Mode
If it doesn’t find your colors, you can see what Excel is really storing. Run:

```
color_mapper.exe input.xlsx config.json --debug
```

It will print lines like:
```
DEBUG: R2C3 -> rgb:FFFF0000
```

Take the code after `->` (e.g., `rgb:FFFF0000`) and put that into your `config.json`.

---

## ✅ Example Workflow
1. You want all **red cells** to say “Check this”.  
2. In `config.json`, write:
   ```json
   {
     "FFFF0000": "Check this"
   }
   ```
3. Run:
   ```
   color_mapper.exe mydata.xlsx config.json
   ```
4. Look in `mydata_output.xlsx` – every row with a red cell now has “Check this” added.

---

## ❓FAQ
**Q: My file is `.xls` not `.xlsx`.**  
👉 Save it in Excel as `.xlsx` first.  

**Q: It didn’t find my color!**  
👉 Use `--debug` to see what code Excel is actually using, then copy that into `config.json`.  

**Q: Do I have to install Python?**  
👉 No. `color_mapper.exe` already includes everything. Just run it.  
