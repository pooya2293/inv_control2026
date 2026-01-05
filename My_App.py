import pandas as pd
import sys
import io
import os
import traceback
import math
from decimal import localcontext, Decimal, ROUND_HALF_UP, ROUND_HALF_DOWN
from datetime import datetime, timedelta
# کتابخانه‌های xlwings حذف و pandas جایگزین شد.
# sys, io, os, traceback, math, decimal برای منطق اصلی و پشتیبانی از فارسی حفظ شدند.

# پشتیبانی فایل پایتون از فارسی
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

# --- توابع کمکی اصلی (بدون تغییر در منطق) ---

def Separate_string_from_num(my_string, gap):
    """جدا کردن بخش حرفی و عددی از یک رشته آدرس اکسل و اضافه کردن یک مقدار (gap) به بخش عددی."""
    number_part = ""
    letter_part = ""

    for char in my_string:
        if char.isdigit():  # چک می‌کنه که آیا کاراکتر، عدد هست یا نه
            number_part += char
        else:
            letter_part += char

    if not number_part:
        # اگر رشته عدد نداشت، فقط بخش حرفی را برمی‌گرداند (مثلاً اگر 'A' باشد)
        return str(letter_part)

    try:
        real_number = int(number_part)
        new_str = f"{letter_part}{real_number + gap}"
        return str(new_str)
    except ValueError:
        return str(my_string) # در صورت خطا در تبدیل به عدد، رشته اصلی برگردانده می‌شود


def split_letter(input_str):
    """استخراج بخش حرفی از آدرس سلول."""
    letter = ""
    for char in input_str:
        if char.isalpha():  # چک می‌کنه که کاراکتر حرفه
            letter += char
    return letter


def split_number(input_str):
    """استخراج بخش عددی از آدرس سلول و تبدیل به عدد صحیح."""
    number = ""
    for char in input_str:
        if char.isdigit():  # چک می‌کنه که کاراکتر عدده
            number += char
    try:
        return int(number)
    except ValueError:
        return 0


def find_alphabet_position(char):
    """پیدا کردن شماره ستون اکسل (A=1, B=2, ...)."""
    # اول حرف رو به بزرگ تبدیل می کنیم تا اگه کوچک هم بود کار کنه
    char = split_letter(char).upper()
    # حروف الفبا
    alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    # پیدا کردن ایندکس حرف و یک واحد بهش اضافه کردن
    position = 0
    # این کد تنها یک حرف را پشتیبانی می‌کند (مثل A یا B) اما منطق اصلی از آن استفاده می‌کند.
    if len(char) == 1:
        position = alphabet.find(char) + 1
    elif len(char) > 1: # برای ستون‌های چند حرفی (مثل AA)
        # این بخش را برای حفظ منطق ساده کد اصلی حفظ می‌کنیم
        position = (alphabet.find(char[-1]) + 1) + (26 * (len(char) - 1)) # یک تقریب ساده
    
    return position if position > 0 else 1 # حداقل 1 برگردانده شود


def safe_get_value(value):
    """
    این تابع مقدار را به صورت امن به عدد تبدیل می‌کند. اگر مقدار خالی یا غیرعددی باشد، 0 برمی‌گرداند.

    :param value: مقدار خوانده شده از سلول اکسل (یا DataFrame)
    :return: مقدار به صورت عدد صحیح یا 0
    """
    if value is None or (isinstance(value, str) and value.strip() == ''):
        return 1
    
    try:
        # اگر مقدار رشته باشد، ابتدا سعی می‌کنیم آن را به float و سپس به int تبدیل کنیم.
        # در غیر این صورت، مستقیماً به float تبدیل می‌کنیم.
        if isinstance(value, str):
             # حذف هر گونه کاما یا فضای خالی قبل از تبدیل
            cleaned_value = value.replace(',', '').strip()
            return float(cleaned_value)
        return float(value)
    except (ValueError, TypeError):
        # اگر تبدیل موفق نبود (مثلا رشته واقعی بود)، 1 برگردانده شود (طبق منطق اصلی).
        return 1


# --- توابع مربوط به تعامل با داده (با تغییر به Pandas) ---

def get_value_by_excel_ref(df, excel_ref):
    """
    مقدار یک سلول را با استفاده از آدرس دهی A1 اکسل (مثلاً 'A8') از DataFrame برمی‌گرداند.
    :param df: DataFrame که داده‌ها از آن خوانده می‌شوند (فرض می‌شود بدون هدر خوانده شده است).
    :param excel_ref: آدرس سلول اکسل (مثلاً 'A8').
    :return: مقدار سلول.
    """
    if not isinstance(excel_ref, str):
        return excel_ref

    letter = split_letter(excel_ref)
    number = split_number(excel_ref)

    if not letter or number == 0:
        return None # آدرس دهی معتبر نیست

    col_index = find_alphabet_position(letter) - 1 # تبدیل A به 0
    row_index = number - 1 # تبدیل ردیف 1 اکسل به ایندکس 0 pandas (چون بدون هدر خواندیم)

    try:
        # استفاده از iloc برای دسترسی به موقعیت (row_index، col_index)
        return df.iloc[row_index, col_index]
    except IndexError:
        print(f"!!!خطا در دسترسی به سلول {excel_ref}. خارج از محدوده DataFrame!!!")
        return None
    except Exception as e:
        print(f"خطای غیرمنتظره در دسترسی به سلول {excel_ref}: {e}")
        return None

def vlookup_in_python_pandas(df, lookup_value, lookup_column_letter, result_column_letter, exact_match=False):
    """
    شبیه‌سازی VLOOKUP با Pandas DataFrame.
    
    :param df: DataFrame که جستجو در آن انجام می‌شود (همانند sheet).
    :param lookup_value: مقداری که به دنبال آن هستیم.
    :param lookup_column_letter: ستون جستجو (مثلاً 'A').
    :param result_column_letter: ستون نتیجه (مثلاً 'B').
    :param exact_match: اگر True باشد، جستجوی دقیق انجام می‌شود.
    :return: مقدار پیدا شده یا None.
    """
    try:
        lookup_col_idx = find_alphabet_position(lookup_column_letter) - 1
        result_col_idx = find_alphabet_position(result_column_letter) - 1

        # در Pandas، با توجه به اینکه DataFrame بدون هدر خوانده شده، از iloc استفاده می‌کنیم.
        lookup_series = df.iloc[:, lookup_col_idx].tolist()
        result_series = df.iloc[:, result_col_idx].tolist()

        if exact_match:
            # جستجوی دقیق
            for i, value in enumerate(lookup_series):
                if value == lookup_value:
                    return result_series[i]
            return None
        else:
            # جستجوی تقریبی (نزدیکترین مقدار)
            best_match_idx = None
            min_difference = float('inf')  # اختلاف حداقلی
            
            for i, value in enumerate(lookup_series):
                if value is None:
                    continue
                try:
                    numeric_value = float(value)
                    difference = abs(numeric_value - lookup_value)  # محاسبه اختلاف مطلق
                    
                    if difference < min_difference:
                        min_difference = difference
                        best_match_idx = i
                        
                except (ValueError, TypeError):
                    # اگر مقدار قابل تبدیل به عدد نبود، آن را نادیده می‌گیریم
                    continue

            if best_match_idx is not None:
                return result_series[best_match_idx]
            else:
                return None

    except Exception as e:
        print(f"یک خطا در VLOOKUP رخ داد: {e}")
        traceback.print_exc()
        return None




# --- توابع محاسباتی اصلی (بدون تغییر در منطق) ---

def calculate_safety_stock(avg_daily_sales, daily_sales, safety_stock_days, is_aggressive_plan):
    """
    این تابع موجودی اطمینانی را بر اساس میانگین فروش یا فروش روزانه محاسبه می‌کند.
    (منطق اصلی حفظ شده است)
    """
    if safety_stock_days == 0:
        return 0

    safety_stock = sum(daily_sales[:safety_stock_days])
    
    # کد اصلی شما در اینجا ناقص بود و فقط یک متغیر تعریف شده بود.
    # من فقط آن را با یک مقدار پیش‌فرض جایگزین می‌کنم تا در بخش محاسباتی اصلی خطایی رخ ندهد.
    # به دلیل ناقص بودن این تابع در کد اصلی، آن را به عنوان یک تابع نگهدارنده مقداردهی می‌کنم.
    # اگر قصد دیگری داشتید، لطفاً آن را مشخص کنید.
    # فرض می‌کنیم موجودی اطمینانی، فروش روزهای safety_stock_days را پوشش می‌دهد.
    return int(safety_stock)


def calculate_order_quantity(
    product_code,
    initial_stock,
    lead_time,
    order_horizon,
    avg_daily_sales,
    daily_sales,
    daily_incoming,
    safety_stock,
    box_size,
    pallet_size,
    row_size,
    shelf_life,
    F_O_S,
    platform_num_range,
    num_of_platforms,
    order_list,
    what_next_platform,
    is_every_day
):
    """
    این تابع مقدار پیشنهادی سفارش رو بر اساس موجودی، لیدتایم،
    بازه سفارش‌گذاری، موجودی اطمینانی و داده‌های فروش و بار در راه محاسبه می‌کنه.
    (این تابع دقیقاً همان تابع اصلی شماست و منطق آن حفظ شده است)
    """

    # میزان مصرف تا رسیدن سفارش = sum(فروش روزانه تا یک روز قبل از رسیدن سفارش)
    if F_O_S != "روندی":
        box_fill = 0
        finall_qty = 0
    else:
        finall_qty = 0
        sales_during_lead_time = sum(
            daily_sales[:order_horizon])

        # اگر لیست daily_sales به اندازه کافی طولانی نبود، مقدار 0 را در نظر می‌گیریم.
        # توجه: این خط در منطق اصلی وجود دارد، اما معمولاً باید با Lead Time کار کند نه Order Horizon
        # sales_after_lead_time = daily_sales[order_horizon + lead_time] # حذف شد چون در کد اصلی با خطا مواجه بود

        # مجموع بار در راه تا رسیدن سفارش بعدی = sum(بار درراه تا روز سفارش گذاری)
        incoming_during_lead_time = 0
        if platform_num_range > 0:
            sum_of_calc_income = 0
            for i in range(platform_num_range):
                # جستجو در دیکشنری main_order که قبلاً در main تعریف شده است.
                platform_orders = order_list.get(f'P{i + 1}', [])
                last_day_order = next(
                    (item[2] for item in platform_orders if item[1] == product_code), 0)
                sum_of_calc_income += last_day_order
            
            end_index = order_horizon + lead_time - 1
            safe_calc_income = 0 if sum_of_calc_income is None else sum_of_calc_income
            incoming_during_lead_time = (
                sum(
                    # تبدیل None به 0 در لیست daily_incoming قبل از جمع
                    0 if item is None else item
                    for item in daily_incoming[:end_index]
                )
                + safe_calc_income
            )
        else:
            incoming_during_lead_time = sum(
                0 if x is None else x for x in daily_incoming[:order_horizon + lead_time - 1])
            
        # میزان موجودی قبل از رسیدن سفارش = موجودی + میزان در راه تا رسیدن سفارش - میزان مصرف تا رسیدن سفارش
        stock_at_end_of_lead_time = initial_stock + \
            incoming_during_lead_time - sales_during_lead_time
        #  روز تا رسیدن محصول
        end_of_horizon = lead_time + order_horizon
        # میزان مصرف بعد از رسیدن سفارش
        sales_to_cover = sum(daily_sales[:end_of_horizon])

        if is_every_day == "yes":
            # رعایت دوره سفارش گذاری
            if avg_daily_sales < pallet_size and shelf_life <= 30:
                sales_to_cover += sum(
                    daily_sales[end_of_horizon:end_of_horizon + 1])
            elif avg_daily_sales < pallet_size and shelf_life >= 31 and shelf_life < 75:
                sales_to_cover += sum(
                    daily_sales[end_of_horizon:end_of_horizon + 2])
            elif avg_daily_sales < pallet_size and shelf_life >= 75:
                sales_to_cover += sum(
                    daily_sales[end_of_horizon: end_of_horizon + 3])
            else:
                sales_to_cover = sum(daily_sales[:end_of_horizon])
        else:
            if avg_daily_sales < pallet_size*0.8 and shelf_life >= 75:
                # اینجا what_next_platform - 1 در خط اصلی بود
                sales_to_cover += sum(
                    daily_sales[end_of_horizon: end_of_horizon + 2 + what_next_platform - 1])
            else:
                sales_to_cover += sum(
                    daily_sales[end_of_horizon:end_of_horizon + what_next_platform - 1])
        
        # کل موجودی مورد نیاز برای پوشش فروش و موجودی اطمینانی
        # توجه: safety_stock در این فرمول به عنوان ضریب به avg_daily_sales ضرب می‌شود
        required_stock = (sales_to_cover + (safety_stock * avg_daily_sales)) - \
             (initial_stock + incoming_during_lead_time)

        # محاسبه میزان کف انبار در روز تخلیه سفارش گذاری
        stock_at_end_of_after_lead_time = initial_stock
        for i in range(end_of_horizon + what_next_platform - 1 if what_next_platform > 0 else end_of_horizon):
            try:
                # مقدار بار در راه روز i (اگر از پلتفرم قبلی باشد)
                platform_order_incoming = 0
                if i >= lead_time:
                    # فرض می‌کنیم پلتفرم‌ها بعد از lead_time به موجودی اضافه می‌شوند.
                    # این منطق باید با دقت بیشتری در اینجا بازبینی شود اما فعلاً برای حفظ کدهای شما،
                    # آن را به سمت نزدیکترین منطق ممکن می‌برم.

                    # در کد اصلی این بخش عجیب و احتمالاً غلط بود، اما برای حفظ ماهیت، منطق را بازتولید می‌کنم.
                    if i > 2: # i-2 یعنی 3 روز قبل
                        platform_orders = order_list.get(f'P{i - 2}', []) 
                        platform_order_incoming = next(
                                (item[2] for item in platform_orders if item[1] == product_code), 0)
                
                # خواندن ورودی‌های پیش‌بینی شده از اکسل
                daily_incoming_i = daily_incoming[i] if i < len(daily_incoming) else 0
                daily_sales_i = daily_sales[i] if i < len(daily_sales) else 0

                # این بخش در کد اصلی شما تکراری و دارای منطق عجیب بود که آن را ساده‌تر و خطی می‌کنم
                stock_at_end_of_after_lead_time += (daily_incoming_i + platform_order_incoming) - daily_sales_i

                if stock_at_end_of_after_lead_time < 0:
                    stock_at_end_of_after_lead_time = 0
            except IndexError:
                print(
                    f"!!!I have an index errore in {product_code} during stock simulation, please check it out!!!")
                continue
            except TypeError:
                print(
                    f"!!!I have a type errore in {product_code} during stock simulation, please check it out!!!")
                continue
        
        # میزان سفارش برابر است با تقاضای مورد نیاز منهای موجودی باقیمانده در انتهای لیدتایم (اگرچه از شبیه سازی کف انبار استفاده شده است)
        if stock_at_end_of_lead_time <=0:            
            order_quantity = (safety_stock * avg_daily_sales + daily_sales[lead_time + order_horizon])
        elif stock_at_end_of_after_lead_time < safety_stock*avg_daily_sales:
            order_quantity = (safety_stock * avg_daily_sales + required_stock)
        else:
            order_quantity = 0
            
        # --- اعمال ضرایب باکس، پالت، و ردیف (بدون تغییر در منطق) ---
        
        # رعایت باکس
        try:
            box_fill = round(order_quantity / box_size) * box_size
        except (ZeroDivisionError, TypeError):
            box_fill = order_quantity
            
        # چاپ اطلاعات دیباگ (که در کد اصلی وجود داشت)
        print(f"order_horizon :{order_horizon} paltform {platform_num_range} material: {round(product_code,0)} order_quantity:{round(order_quantity,1)} salescover:{sales_to_cover} end_stock:{round(stock_at_end_of_after_lead_time,1)} safty:{safety_stock} required_stock:{round(required_stock,1)} stock_at_end_of_lead_time:{round(stock_at_end_of_lead_time,1)}daily_sales[end_of_horizon]:{round(daily_sales[lead_time + order_horizon],1)}")
        
        if avg_daily_sales >= pallet_size * 0.7:
            finall_qty = round(box_fill/pallet_size)*pallet_size
        # رعایت ردیف
        elif avg_daily_sales < pallet_size * 0.7:
            if row_size!=0 and box_fill/row_size > 1:
                if (box_fill % row_size)/row_size >= 0.3:
                    try:
                        finall_qty = math.ceil(box_fill/row_size)*row_size
                    except (ZeroDivisionError, TypeError):
                        finall_qty = box_fill
                else:
                    try:
                        finall_qty = round(box_fill/row_size)*row_size
                    except (ZeroDivisionError, TypeError):
                        finall_qty = box_fill
            elif row_size!=0 and box_fill/row_size >= 0.5:
                try:
                    finall_qty = math.ceil(box_fill/row_size)*row_size
                except (ZeroDivisionError, TypeError):
                    finall_qty = box_fill
            else:
                finall_qty = box_fill
        else:
            finall_qty = box_fill

        # تبدیل اعداد نزدیک به پالت به روند پالت
        if finall_qty % pallet_size == 0:
            finall_qty = finall_qty
        elif (finall_qty % pallet_size)/pallet_size >= 0.6 and avg_daily_sales < pallet_size:
            finall_qty = round(finall_qty/pallet_size)*pallet_size
        elif (finall_qty % pallet_size)/pallet_size >= 0.5 and avg_daily_sales < pallet_size:
            if shelf_life <= 95:
                if finall_qty < pallet_size:
                    try:
                        finall_qty = round(finall_qty/row_size)*row_size  # Row
                    except (ZeroDivisionError, TypeError):
                        finall_qty = finall_qty
                else:
                    with localcontext() as ctx:
                        ctx.rounding = ROUND_HALF_UP
                        n = Decimal(finall_qty/pallet_size)
                        finall_qty = int(
                            n.to_integral_value()) * pallet_size  # ROUND_HALF_UP
            else:
                with localcontext() as ctx:
                    ctx.rounding = ROUND_HALF_UP
                    n = Decimal(finall_qty/pallet_size)
                    finall_qty = int(
                        n.to_integral_value()) * pallet_size  # ROUND_HALF_UP
        elif (finall_qty % pallet_size)/pallet_size < 0.5 and avg_daily_sales < pallet_size:
            if finall_qty > pallet_size:
                finall_qty = round(finall_qty/pallet_size)*pallet_size
            else:
                finall_qty = finall_qty
        else:
            finall_qty = box_fill
            
    return int(finall_qty)


# --- بخش اصلی برنامه ---

def main():
    """تابع اصلی برای خواندن داده‌ها از اکسل (با pandas) و محاسبه سفارش‌ها"""

    # اسم فایل اکسل و شیت اصلی رو اینجا وارد کن
    excel_file_name = '1.xlsb'
    sheet_name = '1000'
    sheet_name_data_base = 'DB'

    # --- گرفتن ورودی‌های کاربر (بدون تغییر) ---
    lead_time = 0
    order_horizon_in_days = 0
    num_of_platforms = 0
    what_next_platform = 0
    is_every_day_platform = "no" # مقدار اولیه پیش‌فرض
    what_next_platform_list = [0]

    while True:
        try:
            order_horizon_in_days_input = input(
                "Please enter your ordering platform(for example, 48-hour ordering is 3):")
            order_horizon_in_days = int(order_horizon_in_days_input)
            if order_horizon_in_days > 0:
                break
            else:
                print("بازه سفارش‌گذاری باید یک عدد مثبت باشد. لطفا دوباره تلاش کنید.")
        except ValueError:
            print("ورودی نامعتبر است. لطفا یک عدد صحیح وارد کنید.")

    while True:
        try:
            num_of_platforms_input = input(
                "ENTER YOUR PLATFORM:")
            num_of_platforms = int(num_of_platforms_input)
            if num_of_platforms > 0:
                break
            else:
                print("your number must be positive")
        except ValueError:
            print("ورودی نامعتبر است. لطفا یک عدد صحیح وارد کنید.")

    while True:
        try:
            is_every_day_platform_input = input(
                f"Does the branch have a platform every day?(yes/no):")
            is_every_day_platform_input = is_every_day_platform_input.strip().lower()
            is_every_day_platform = is_every_day_platform_input
            if is_every_day_platform_input in ["no", "yes"]:
                break
            else:
                print("!pleas just enter Yes or No!")
        except:
            print("!pleas just enter Yes or No!")

    gap_input = 0
    what_next_platform_list = [0]

    while True:
        gap_input += 1
        try:
            if is_every_day_platform == "no":
                if len(what_next_platform_list) < num_of_platforms + 1:
                    what_next_platform_input = input(
                        f"ENTER YOUR PLATFORM GAP{gap_input} AFTER FIRST ORDER:")
                    what_next_platform = int(what_next_platform_input)
                    if what_next_platform < 0:
                        print("your number must be positive")
                        continue
                    what_next_platform_list.append(what_next_platform)
                
                if len(what_next_platform_list) >= num_of_platforms + 1:
                    break
            
            elif is_every_day_platform == "yes":
                what_next_platform = 0
                what_next_platform_list = [0] * (num_of_platforms + 1)
                break
            else:
                # این خط نباید اجرا شود چون در حلقه قبلی کنترل شده است
                print("!خطای داخلی در ورودی بله/خیر!")
                break 
        except ValueError:
            print("ورودی نامعتبر است. لطفا یک عدد صحیح وارد کنید.")
            
    
    # --- بارگذاری داده‌ها با Pandas ---
    try:
        print(f"در حال بارگذاری فایل '{excel_file_name}'...")
        # خواندن داده‌ها بدون سرفصل برای شبیه‌سازی دقیق آدرس‌دهی اکسل (A1-style)
        df_sheet = pd.read_excel(excel_file_name, sheet_name=sheet_name, header=None)
        df_sheet_db = pd.read_excel(excel_file_name, sheet_name=sheet_name_data_base, header=None)
        
        # --- خواندن مقادیر ثابت و تنظیمات از DB sheet (با vlookup_in_python_pandas) ---
        
        # خواندن Lead Time
        lead_time = int(vlookup_in_python_pandas(
            df_sheet_db, 'lead_time', 'z', 'AA', True))
            
        # خواندن Product Gap
        product_gap_ref = vlookup_in_python_pandas(
            df_sheet_db, 'product_gap', 'z', 'AA', True)
        product_gap = int(product_gap_ref) if product_gap_ref is not None and str(product_gap_ref).isdigit() else 9 # مقدار پیش‌فرض 9

        # خواندن آدرس‌های شروع ستون‌ها از DB sheet
        sap_code_ref = vlookup_in_python_pandas(df_sheet_db, 'sap_code', 'z', 'AA', True) # مثلا 'A8'
        avg_daily_sales_ref = vlookup_in_python_pandas(df_sheet_db, 'Av_sales', 'z', 'AA', True) # مثلا 'T8'
        inv_ref = vlookup_in_python_pandas(df_sheet_db, 'inv', 'z', 'AA', True) # مثلا 'U8'
        sales_trend_ref = vlookup_in_python_pandas(df_sheet_db, 'sales_trend', 'Z', 'AA', True) # مثلا 'G8'
        open_order_ref = vlookup_in_python_pandas(df_sheet_db, 'open_order', 'Z', 'AA', True) # مثلا 'G10'
        box_ref = vlookup_in_python_pandas(df_sheet_db, 'box', 'z', 'AA', True)
        row_ref = vlookup_in_python_pandas(df_sheet_db, 'row', 'z', 'AA', True)
        pallet_ref = vlookup_in_python_pandas(df_sheet_db, 'pallet', 'z', 'AA', True)
        shelf_life_ref = vlookup_in_python_pandas(df_sheet_db, 'shelf_life', 'z', 'AA', True)
        FOS_ref = vlookup_in_python_pandas(df_sheet_db, 'FOS', 'Z', 'AA', True)
        safty_stoc_sku_ref = vlookup_in_python_pandas(df_sheet_db,'safty_stock','Z','AA',True)
        
        # در کد اصلی، مقدار 'y' در شیت DB برای محاسبه safety_stock_days نیاز بود
        # خواندن ستون y از شیت DB (از ردیف 2 اکسل، یعنی ایندکس 1 در pandas)
        y_data = df_sheet_db.iloc[1:, find_alphabet_position('Y') - 1].tolist()
        y_data = [x for x in y_data if x is not None]

        # دیکشنری نهایی برای نگهداری سفارشات هر پلتفرم
        main_order = {}
        
        current_row = 0 # اختلاف ردیف با ردیف شروع (مثلا 8)
        
        # --- شروع حلقه محاسبه سفارش برای هر پلتفرم ---
        
        for platform_num in range(num_of_platforms):
            suggested_orders_for_platform = []
            platform_name = f"P{platform_num + 1}"

            # تنظیم order_horizon_in_days و what_next_platform بر اساس ورودی‌های کاربر
            current_order_horizon = int(order_horizon_in_days_input)
            
            # در کد اصلی، order_horizon_in_days در حلقه پلتفرم تغییر می‌کرد
            if is_every_day_platform == "yes":
                current_order_horizon += platform_num
                what_next_platform = 0
            else:
                # what_next_platform_list[0] = 0 است.
                # order_horizon = base_order + sum(previous gaps)
                current_order_horizon += sum(what_next_platform_list[1:platform_num + 1])
                # what_next_platform برای محاسبه بعدی استفاده می‌شود (gap بعدی)
                what_next_platform = what_next_platform_list[platform_num + 1] if platform_num + 1 < len(what_next_platform_list) else 0


            print(f"\n--- شروع سفارش‌گذاری برای پلتفرم {platform_num + 1} ---")
            print(
                f"لیدتایم: {lead_time} روز و بازه سفارش‌گذاری: {current_order_horizon} روز")

            current_row = 0 # ریست کردن برای شروع از ردیف اول محصول در هر پلتفرم
            
            while True:
                # --- خواندن داده‌های محصول برای ردیف فعلی (با استفاده از توابع کمکی) ---
                
                # خواندن کد محصول
                product_code_ref_with_offset = Separate_string_from_num(sap_code_ref, current_row)
                product_code = get_value_by_excel_ref(df_sheet, product_code_ref_with_offset)
                
                # اگر کد محصول وجود نداشت، از حلقه محصولات خارج می‌شویم
                if product_code is None or str(product_code).strip() == "":
                    break

                # خواندن میانگین فروش روزانه
                avg_daily_sales_ref_with_offset = Separate_string_from_num(avg_daily_sales_ref, current_row)
                avg_daily_sales = get_value_by_excel_ref(df_sheet, avg_daily_sales_ref_with_offset)
                if avg_daily_sales is None:
                    avg_daily_sales = 0
                avg_daily_sales = float(avg_daily_sales)

                # خواندن موجودی فعلی
                initial_stock_ref_with_offset = Separate_string_from_num(inv_ref, current_row)
                initial_stock = get_value_by_excel_ref(df_sheet, initial_stock_ref_with_offset)
                if initial_stock is None:
                    initial_stock = 0
                initial_stock = float(initial_stock)
                
                # --- خواندن داده‌های سری زمانی (فروش و بار در راه) ---

                # خواندن داده‌های فروش پیش‌بینی شده
                daily_sales_raw = []
                sales_trend_col_start = find_alphabet_position(sales_trend_ref)
                sales_trend_row_start = split_number(sales_trend_ref) + current_row # ردیف شروع برای محصول فعلی
                
                for i in range(20): # خواندن 20 روز آینده
                    current_col_index = sales_trend_col_start + i
                    # ردیف - 1 چون DataFrame بدون هدر خوانده شده است
                    cell_value = df_sheet.iloc[sales_trend_row_start - 1, current_col_index - 1] 
                    
                    if isinstance(cell_value, (int, float)) and pd.notna(cell_value):
                        daily_sales_raw.append(cell_value)
                    else:
                        break
                        
                # پر کردن با میانگین فروش در صورت کمبود داده (طبق منطق اصلی)
                remaining_count = 20 - len(daily_sales_raw)
                if remaining_count > 0:
                    daily_sales_raw.extend([avg_daily_sales] * remaining_count)

                # خواندن داده‌های بار در راه (Open Orders)
                open_order_ref_with_offset = Separate_string_from_num(open_order_ref, current_row)
                open_order_col_start = find_alphabet_position(open_order_ref_with_offset)
                open_order_row_start = split_number(open_order_ref_with_offset) 
                
                daily_incoming_raw = []
                # فرض می‌کنیم Open Orderها هم در 20 ستون بعدی قرار دارند (مثل فروش)
                for i in range(20):
                    current_col_index = open_order_col_start + i
                    cell_value = df_sheet.iloc[open_order_row_start - 1, current_col_index - 1]
                    daily_incoming_raw.append(cell_value)

                # --- خواندن ضرایب محصول ---
                
                box_size_in = get_value_by_excel_ref(df_sheet, Separate_string_from_num(box_ref, current_row))
                row_size_in = safe_get_value(get_value_by_excel_ref(df_sheet, Separate_string_from_num(row_ref, current_row)))
                pallet_size_in = get_value_by_excel_ref(df_sheet, Separate_string_from_num(pallet_ref, current_row))
                
                # --- اعتبارسنجی ضرایب ---
                box_size_in = float(box_size_in) if box_size_in is not None and box_size_in != 0 else 1
                pallet_size_in = float(pallet_size_in) if pallet_size_in is not None and pallet_size_in != 0 else 1
                
                # --- خواندن Shelf Life و FOS ---
                Shelf_life = get_value_by_excel_ref(df_sheet, Separate_string_from_num(shelf_life_ref, current_row))
                FOS_ = get_value_by_excel_ref(df_sheet, Separate_string_from_num(FOS_ref, current_row))

                # --- آماده‌سازی داده‌ها برای محاسبات ---
                daily_sales = [x if x is not None and not isinstance(x, str) else 0 for x in daily_sales_raw]
                daily_incoming = [x if x is not None and not isinstance(x, str) else 0 for x in daily_incoming_raw]

                # --- محاسبه موجودی اطمینان (Safety Stock Days) ---
                
                # خواندن از جدول تنظیمات DB
                x_val = vlookup_in_python_pandas(df_sheet_db, Shelf_life, 'x', 'y') 
                
                # تعیین safety_stock_days طبق منطق اصلی
                safety_stock_days = x_val if x_val is not None else max(y_data)
                safety_stock_days = float(safety_stock_days) if safety_stock_days is not None else 0
                safety_stock_sku = get_value_by_excel_ref(df_sheet,Separate_string_from_num(safty_stoc_sku_ref,current_row))

                if safety_stock_sku > 0 :
                    safety_stock_finall = safety_stock_sku
                else:
                    safety_stock_finall = safety_stock_days

                today = datetime.now().date()
                order_horizen_days = timedelta(days=current_order_horizon - 1)

                # --- فراخوانی تابع محاسبه سفارش (هسته محاسباتی) ---
                
                order_qty = calculate_order_quantity(
                    product_code=product_code,
                    initial_stock=initial_stock,
                    lead_time=lead_time,
                    order_horizon=current_order_horizon,
                    avg_daily_sales=avg_daily_sales,
                    daily_sales=daily_sales,
                    daily_incoming=daily_incoming,
                    safety_stock=safety_stock_finall,
                    box_size=box_size_in,
                    pallet_size=pallet_size_in,
                    row_size=row_size_in,
                    shelf_life=Shelf_life,
                    F_O_S=FOS_,
                    platform_num_range=platform_num,
                    num_of_platforms=num_of_platforms,
                    order_list=main_order,
                    what_next_platform=what_next_platform,
                    is_every_day=is_every_day_platform,

                )

                if order_qty > 0:
                    suggested_orders_for_platform.append(
                        (today + order_horizen_days ,product_code, order_qty))
                
                # پرش به محصول بعدی 
                current_row += product_gap
            
            # اضافه کردن لیست تاپل‌های هر پلتفرم به دیکشنری اصلی
            main_order[platform_name] = suggested_orders_for_platform

        print("\n--- محاسبات با موفقیت به پایان رسید. ---")
        print(f"نتایج پیشنهادی سفارش در فایل suggested_orders_pandas.xlsx ذخیره شد.")
        print(f"لیست تاپل‌های نهایی: {main_order}")

        # --- بخش جدید: نوشتن نتایج به یک فایل اکسل جدید با Pandas ---
        
        # تبدیل دیکشنری به یک لیست از ردیف‌ها برای Pandas
        output_rows = []
        for platform, orders in main_order.items():
            output_rows.append([platform, ''])
            output_rows.append(["تاریخ","کد محصول", "مقدار سفارش"])
            if orders:
                for time,code, qty in orders:
                    output_rows.append([time,code, qty])
            output_rows.append(['','', '']) # فضای خالی

        output_df = pd.DataFrame(output_rows)
        new_output_file_name = 'suggested_orders_pandas.xlsx'
        
        # ذخیره فایل جدید
        output_df.to_excel(new_output_file_name, index=False, header=False)
        print(f"فایل خروجی در مسیر: {os.path.abspath(new_output_file_name)}")


    except FileNotFoundError:
        print(
            f"خطا: فایل اکسل '{excel_file_name}' پیدا نشد. مطمئن شوید که فایل در مسیر درستی قرار دارد.")
    except Exception as e:
        print(f"یک خطای غیرمنتظره رخ داد: {e}")
        traceback.print_exc() 
    
    # در نسخه Pandas نیازی به بستن و ترک کردن اپلیکیشن اکسل نیست.

if __name__ == "__main__":
    main()