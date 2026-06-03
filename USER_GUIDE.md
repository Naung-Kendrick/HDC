# HDC - User Guidelines
## အသုံးပြုသူလက်စွဲ

---

## Before You Upload - စစ်ဆေးရန်အချက်များ

### ✅ DO - လုပ်ဆောင်ရန်

| Guidelines | အကြံပြုချက်များ |
|------------|----------------|
| **Use Unicode Myanmar Font** | ယူနီကုဒ် မြန်မာဖောင့် အသုံးပြုရန် |
| **Fill All Required Fields** | လိုအပ်သော အချက်အလက်ကွက်လပ်များအားလုံး ဖြည့်စွက်ရန် |
| **Check Spelling Before Upload** | စာမတင်မီ စာလုံးပေါင်းသတ်ပုံကို စစ်ဆေးရန် |
| **Use DD-MM-YYYY Date Format** | ရက်စွဲပုံစံကို ရက်-လ-ခုနှစ် (DD-MM-YYYY) အတိုင်း အသုံးပြုရန် |
| **Verify Household Numbers** | အိမ်ထောင်စုစာရင်း နံပါတ်များကို မှန်ကန်မှု ရှိ၊ မရှိ စစ်ဆေးရန် |

### ❌ DON'T - ရှောင်ရန်

| Don't | ရှောင်ရန် |
|-------|-----------|
| **Leave Required Fields Empty** | လိုအပ်သော အချက်အလက်ကွက်လပ်များကို ဗလာ (အလွတ်) မထားရ |
| **Mix Zawgyi & Unicode Fonts** | ဇော်ဂျီနှင့် ယူနီကုဒ် ဖောင့်များကို ရောနှောမသုံးရ |
| **Use Special Characters** | အထူးပြုလုပ်ထားသော သင်္ကေတ/စာလုံးများကို မသုံးရ |
| **Merge Cells in Excel** | Excel တွင် ကွက်လပ် (Cells) များကို ပေါင်းစပ်ခြင်း မပြုရ |
| **Add Duplicate Entries** | ဒေတာအချက်အလက်များကို ထပ်ခါတလဲလဲ (နှစ်ခါ) မထည့်ရ |

---

## Required Fields - မဖြစ်မနေဖြည့်သွင်းရန်

| Field | အချက်အလက် |
|-------|-----------|
| **Ward/Village/Group** | ရပ်ကွက် / ကျေးရွာအုပ်စု / ကျေးရွာ |
| **Township** | မြို့နယ် |
| **District** | ခရိုင် |
| **Gender** | ကျား/မ (ကျား၊ မ ရွေးချယ်ရန်) |
| **Relationship** | တော်စပ်ပုံ |
| **Name (Myanmar)** ⭐ | အမည် (မြန်မာဘာသာ) - Recommended |

🔴 **Red dot** = Required (မဖြစ်မနေလိုအပ်)  
🟠 **Orange dot** = Recommended (အကြံပြုထားသော)

---

## How to Use - အသုံးပြုနည်း

### Step 1: Upload Your File
1. Click the upload area OR drag & drop your file
2. Supported formats: **.XLSX, .XLS, .CSV**
3. Max size: **10MB**
4. First row must contain **headers** (column names)

### Step 2: Wait for Validation
- Processing takes 1-3 seconds
- You'll see a loading spinner
- System checks all data automatically

### Step 3: Review Results

#### ✅ If "All Checks Passed"
- File is clean and ready
- Click **"Download Corrected CSV"**
- Use this CSV for database import

#### ⚠️ If Errors Found
- Click **"Download Error Report"** (.xlsx file)
- Open the Excel file
- Check **"Errors"** sheet - fix these first
- Check **"Warnings"** sheet - review if needed
- Fix issues in your original file
- Re-upload for validation

### Step 4: Download & Continue
- Download corrected CSV
- Or click **"Check Another File"** to validate more

---

## Common Issues - ပုံမှန်ပြဿနာများ

### Issue 1: "File size exceeds 10MB"
**Solution**: Split into smaller files or remove unnecessary columns

### Issue 2: "N/A Error" for Names
**Solution**: 
- Check for merged cells in Excel
- Ensure Name column has data
- Don't leave cells empty

### Issue 3: Myanmar text showing boxes □□□
**Solution**: Use Unicode font (Padauk, Myanmar3) - NOT Zawgyi

### Issue 4: Date rejected
**Solution**: Use **DD-MM-YYYY** format (e.g., 15-06-1990)

### Issue 5: "Household number already exists"
**Solution**: Check for duplicate entries in your data

---

## Understanding Error Report - အမှားအယွင်းအစီရင်ခံစာ

Your downloaded `.xlsx` file contains:

| Sheet | Content | အကြောင်းအရာ |
|-------|---------|-------------|
| **Summary** | Overview statistics | စုစုပေါင်း အနှစ်ချုပ် |
| **Errors** | Must-fix issues | မဖြစ်မနေပြင်ရမည့်အချက်များ |
| **Warnings** | Review recommended | စစ်ဆေးသင့်သည့်အချက်များ |

### Error Report Columns:
- **Row Number** - မှားနေသော row အမှတ်
- **Name** - အမည်
- **Missing Fields** - မဖြည့်ထားသောအချက်များ
- **Myanmar Text Issues** - မြန်မာစာပြဿနာများ

---

## File Requirements - ဖိုင် လိုအပ်ချက်များ

✅ **Supported Formats**: .XLSX, .XLS, .CSV  
✅ **Max Size**: 10MB  
✅ **Headers**: First row must have column names  
✅ **Encoding**: UTF-8 (for Myanmar text support)

---

## Quick Tips - အထူးအကြံပြုချက်များ

1. **Always validate before database import**
2. **Keep a backup** of your original file
3. **Fix errors first**, then warnings
4. **Re-upload after fixing** to confirm all clear
5. **Use the corrected CSV** for final import

---

## Need Help? - အကူအညီလိုပါက

**Developed by**: Mai Naung Naung & Mai Nay Lin  
**Version**: v1.0  
**Website**: https://github.com/Naung-Kendrick/HDC

---

**Happy Validating!** 🎉
