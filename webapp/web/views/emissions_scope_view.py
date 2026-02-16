from flask import Blueprint, render_template, request, jsonify, url_for
from flask_login import login_required, current_user
from ...models import FormAndFormula, Scope, Material, CampusAndDepartment
from ...models.materail_model import QuantityType
from datetime import datetime
from ..utils.acl import permissions_required_all
from flask import make_response
import json
import urllib.parse
import openpyxl
from ..views.emissoins import save_material

module = Blueprint("emissions_scope", __name__, url_prefix="/emissions-scope")


@module.route("/", methods=["GET"])
@login_required
@permissions_required_all(["เข้าถึงหน้าข้อมูลการปล่อย"])
def emissions_scope():
    # รับปีที่เลือกจาก query parameter หรือใช้ปีปัจจุบันเป็นค่าเริ่มต้น
    selected_year = request.args.get("year", default=datetime.now().year, type=int)
    user = current_user
    user.campus = CampusAndDepartment.get_campus_name(user.campus_id)
    user.department = CampusAndDepartment.get_department_name(
        user.campus_id, user.department_key
    )

    # ดึงข้อมูล Scope ที่ตรงกับ campus และ department ของ current_user
    # และกรองเฉพาะ scope ที่อยู่ในฟิลด์ของ user เท่านั้น
    scopes = []

    # ดึง Scope 1 ที่อยู่ในฟิลด์ ghg_scope_1 ของ user
    if user.ghg_scope_1:
        scope_1_list = Scope.objects(
            ghg_scope=1,
            ghg_sup_scope__in=user.ghg_scope_1,
            campus=user.campus_id,
            department=user.department_key,
        )
        scopes.extend(scope_1_list)

    # ดึง Scope 2 ที่อยู่ในฟิลด์ ghg_scope_2 ของ user
    if user.ghg_scope_2:
        scope_2_list = Scope.objects(
            ghg_scope=2,
            ghg_sup_scope__in=user.ghg_scope_2,
            campus=user.campus_id,
            department=user.department_key,
        )
        scopes.extend(scope_2_list)

    # ดึง Scope 3 ที่อยู่ในฟิลด์ ghg_scope_3 ของ user
    if user.ghg_scope_3:
        scope_3_list = Scope.objects(
            ghg_scope=3,
            ghg_sup_scope__in=user.ghg_scope_3,
            campus=user.campus_id,
            department=user.department_key,
        )
        scopes.extend(scope_3_list)

    # ดึงปีทั้งหมดที่มีในฐานข้อมูล Material สำหรับ dropdown
    all_years = Material.objects().distinct("year")
    all_years = sorted([year for year in all_years if year is not None], reverse=True)

    # ถ้าไม่มีปีในฐานข้อมูล ให้เพิ่มปีปัจจุบัน
    if not all_years:
        all_years = [datetime.now().year]
    elif selected_year not in all_years:
        all_years.append(selected_year)
        all_years.sort(reverse=True)

    # จัดกลุ่ม Scope ตาม ghg_scope
    grouped_scopes = {}
    overall_progress = 0
    total_sources = len(scopes)
    completed = 0
    in_progress = 0
    not_started = 0

    for scope in scopes:
        main_scope = f"Scope {scope.ghg_scope}"
        if main_scope not in grouped_scopes:
            grouped_scopes[main_scope] = []

        # คำนวณ Progress สำหรับ Scope นี้ (เฉพาะปีที่เลือก)
        progress = calculate_scope_progress(scope, selected_year)

        # กำหนดสถานะตาม Progress
        if progress == 100:
            status = "Completed"
            completed += 1
        elif progress > 0:
            status = "In progress"
            in_progress += 1
        else:
            status = "Not started"
            not_started += 1

        # เพิ่ม Progress รวม
        overall_progress += progress

        # เพิ่ม progress และ status เข้าไปใน scope object
        scope.progress = round(progress, 1)
        scope.status = status

        # ส่ง scope object ทั้งหมดไปเลย
        grouped_scopes[main_scope].append(scope)

    # คำนวณ Overall Progress
    overall_progress = (
        round(overall_progress / total_sources, 1) if total_sources > 0 else 0
    )

    return render_template(
        "/emissions-scope/emissions-scope.html",
        user=user,
        mockup_data={
            "overall_progress": overall_progress,
            "total_sources": total_sources,
            "in_progress": in_progress,
            "not_started": not_started,
            "completed": completed,
            "scopes": grouped_scopes,
            "selected_year": selected_year,
            "all_years": all_years,
        },
    )


def calculate_scope_progress(scope, selected_year):
    """
    คำนวณ Progress:
    - นับเฉพาะ Material ที่อยู่ใน scope นี้เท่านั้น
    - นับ Material ที่มี quantity_type (ไม่เป็น None และไม่ว่าง)
    - Material ที่ถูกลบจะไม่มี field quantity_type หรือเป็น None/[] จะไม่ถูกนับ
    """
    num_head_table = len(scope.head_table)
    if num_head_table == 0:
        return 0

    total_fields_required = num_head_table * 12
    if total_fields_required == 0:
        return 0

    materials_qs = Material.objects(
        scope=scope.ghg_scope,
        sub_scope=scope.ghg_sup_scope,
        year=selected_year,
        campus=current_user.campus_id,
        department=current_user.department_key,
    )

    # นับเฉพาะ Material ที่มี quantity_type และไม่เป็น None และไม่ว่าง
    filled = materials_qs.filter(
        quantity_type__exists=True,  # มี field quantity_type
        quantity_type__ne=None,  # ไม่เป็น None
        quantity_type__not__size=0,  # ไม่ว่าง (มีอย่างน้อย 1 รายการ)
    ).count()

    progress = (filled / total_fields_required) * 100
    return min(progress, 100)


@module.route("/get-latest-sub-scope", methods=["POST"])
@login_required
@permissions_required_all(["แก้ไขข้อมูลการปล่อย"])
def get_latest_sub_scope():
    ghg_scope = request.json.get("ghg_scope")  # รับข้อมูลจาก JSON
    if not ghg_scope or not ghg_scope.isdigit():
        return jsonify({"latest_sub_scope": 1})  # ถ้า Scope หลักว่างหรือไม่ใช่ตัวเลข ให้เริ่มที่ 1

    ghg_scope = int(ghg_scope)

    # แก้ไข: เพิ่มการกรองตาม campus และ department ของ current_user
    latest_sub_scope = (
        Scope.objects(
            ghg_scope=ghg_scope,
            campus=current_user.campus,
            department=current_user.department,
        )
        .order_by("-ghg_sup_scope")
        .first()
    )

    latest_sub_scope = latest_sub_scope.ghg_sup_scope + 1 if latest_sub_scope else 1
    return jsonify({"latest_sub_scope": latest_sub_scope})


@module.route("/edit/<int:ghg_scope>/<int:ghg_sup_scope>", methods=["GET", "POST"])
@login_required
@permissions_required_all(["แก้ไขข้อมูลการปล่อย"])
def edit_scope(ghg_scope, ghg_sup_scope):
    # ค้นหา Scope ที่ต้องการแก้ไขตาม campus และ department ของ current_user
    scope = Scope.objects(
        ghg_scope=ghg_scope,
        ghg_sup_scope=ghg_sup_scope,
        campus=current_user.campus_id,
        department=current_user.department_key,
    ).first()

    if not scope:
        return render_template(
            "/emissions-scope/edit-scope-error.html",
            error="ไม่พบ Scope ที่ต้องการแก้ไข หรือคุณไม่มีสิทธิ์เข้าถึง",
        )

    if request.method == "POST":
        head_table = request.form.getlist("head_table")

        # อัปเดตข้อมูล (ไม่ต้องตรวจสอบ name และ desc เพราะเป็น hidden input)
        scope.head_table = head_table if head_table else []
        scope.save()

        # ใช้ toast notification แทน popup

        response = make_response("")
        encoded_message = urllib.parse.quote("แก้ไข Scope สำเร็จ!")

        trigger_data = {
            "closeModal": True,
            "showSuccess": encoded_message,
            "refreshPage": True,
        }
        response.headers["HX-Trigger"] = json.dumps(trigger_data)

        return response

    # กรณี GET: แสดงฟอร์มแก้ไข
    # ดึง material_names จาก FormAndFormula ที่ตรงกับ scope และ sub_scope
    materials = FormAndFormula.objects(
        ghg_scope=ghg_scope,
        ghg_sup_scope=ghg_sup_scope,  # แก้จาก ghg_sub_scope เป็น ghg_sup_scope
    ).distinct("material_name")

    material_names = sorted([name for name in materials if name])

    # ดึงชื่อจริงของ campus และ department
    campus_name = ""
    department_name = ""

    try:
        from ...models.campus_and_department_model import CampusAndDepartment

        campus_doc = CampusAndDepartment.objects(id=current_user.campus_id).first()
        if campus_doc:
            campus_name = campus_doc.name.get("0", "")
            department_name = campus_doc.departments.get(
                current_user.department_key, ""
            )
    except Exception as e:
        print(f"Error getting campus/department names: {e}")
        campus_name = scope.campus
        department_name = scope.department

    return render_template(
        "/emissions-scope/edit-scope.html",
        scope=scope,
        material_names=material_names,
        campus_name=campus_name,
        department_name=department_name,
    )


@module.route("/scope-description/<int:ghg_scope>/<int:ghg_sup_scope>", methods=["GET"])
@login_required
def scope_description(ghg_scope, ghg_sup_scope):
    """แสดง popup description ของ Scope"""
    try:
        # ค้นหา Scope ที่ต้องการ
        scope = Scope.objects(
            ghg_scope=ghg_scope,
            ghg_sup_scope=ghg_sup_scope,
            campus=current_user.campus_id,
            department=current_user.department_key,
        ).first()

        if not scope:
            return render_template(
                "/emissions-scope/partials/scope-description-modal.html",
                error="ไม่พบข้อมูล Scope ที่ต้องการ",
            )

        # ดึงชื่อจริงของ campus และ department
        campus_name = ""
        department_name = ""

        try:
            from ...models.campus_and_department_model import CampusAndDepartment

            campus_doc = CampusAndDepartment.objects(id=current_user.campus_id).first()
            if campus_doc:
                campus_name = campus_doc.name.get("0", "")
                department_name = campus_doc.departments.get(
                    current_user.department_key, ""
                )
        except Exception as e:
            print(f"Error getting campus/department names: {e}")
            campus_name = scope.campus
            department_name = scope.department

        # ดึงผู้รับผิดชอบ subscope นี้ (username, name, email)
        from ...models import User

        responsibles = []
        if scope.ghg_scope == 1:
            users = User.objects(
                campus_id=current_user.campus_id,
                department_key=current_user.department_key,
                ghg_scope_1=scope.ghg_sup_scope,
            )
        elif scope.ghg_scope == 2:
            users = User.objects(
                campus_id=current_user.campus_id,
                department_key=current_user.department_key,
                ghg_scope_2=scope.ghg_sup_scope,
            )
        elif scope.ghg_scope == 3:
            users = User.objects(
                campus_id=current_user.campus_id,
                department_key=current_user.department_key,
                ghg_scope_3=scope.ghg_sup_scope,
            )
        else:
            users = []

        responsibles = [
            {"username": u.username, "name": u.name, "email": u.email} for u in users
        ]
        return render_template(
            "/emissions-scope/partials/scope-description-modal.html",
            scope=scope,
            campus_name=campus_name,
            department_name=department_name,
            responsibles=responsibles,
        )

    except Exception as e:
        return render_template(
            "/emissions-scope/partials/scope-description-modal.html",
            error=f"เกิดข้อผิดพลาด: {str(e)}",
        )


@module.route("/load-import-modal", methods=["GET"])
@login_required
@permissions_required_all(["แก้ไขข้อมูลการปล่อย"])
def load_import_modal():
    """แสดง modal สำหรับ import Excel"""
    # ดึงข้อมูลปีที่มีอยู่ในระบบ
    all_years = Material.objects().distinct("year")
    all_years = sorted([year for year in all_years if year is not None], reverse=True)

    # ถ้าไม่มีปีในฐานข้อมูล ให้เพิ่มปีปัจจุบัน
    if not all_years:
        all_years = [datetime.now().year]

    # ดึง Scope ทั้งหมดของ user
    scopes = []

    # ดึง Scope 1 ที่อยู่ในฟิลด์ ghg_scope_1 ของ user
    if current_user.ghg_scope_1:
        scope_1_list = Scope.objects(
            ghg_scope=1,
            ghg_sup_scope__in=current_user.ghg_scope_1,
            campus=current_user.campus_id,
            department=current_user.department_key,
        )
        scopes.extend(
            [
                {
                    "ghg_scope": s.ghg_scope,
                    "ghg_sup_scope": s.ghg_sup_scope,
                    "ghg_name": s.ghg_name,
                }
                for s in scope_1_list
            ]
        )

    # ดึง Scope 2 ที่อยู่ในฟิลด์ ghg_scope_2 ของ user
    if current_user.ghg_scope_2:
        scope_2_list = Scope.objects(
            ghg_scope=2,
            ghg_sup_scope__in=current_user.ghg_scope_2,
            campus=current_user.campus_id,
            department=current_user.department_key,
        )
        scopes.extend(
            [
                {
                    "ghg_scope": s.ghg_scope,
                    "ghg_sup_scope": s.ghg_sup_scope,
                    "ghg_name": s.ghg_name,
                }
                for s in scope_2_list
            ]
        )

    # ดึง Scope 3 ที่อยู่ในฟิลด์ ghg_scope_3 ของ user
    if current_user.ghg_scope_3:
        scope_3_list = Scope.objects(
            ghg_scope=3,
            ghg_sup_scope__in=current_user.ghg_scope_3,
            campus=current_user.campus_id,
            department=current_user.department_key,
        )
        scopes.extend(
            [
                {
                    "ghg_scope": s.ghg_scope,
                    "ghg_sup_scope": s.ghg_sup_scope,
                    "ghg_name": s.ghg_name,
                }
                for s in scope_3_list
            ]
        )

    return render_template(
        "/emissions-scope/partials/import-excel-modal.html",
        years=all_years,
        scopes=scopes,
        current_year=datetime.now().year,
    )


@module.route("/import-excel", methods=["POST"])
@login_required
@permissions_required_all(["แก้ไขข้อมูลการปล่อย"])
def import_excel():
    """
    อัปโหลดและประมวลผลไฟล์ Excel สำหรับ import ข้อมูล Material
    """
    try:
        # รับค่าจากฟอร์ม
        scope_id = request.form.get("scope_id")
        sub_scope_id = request.form.get("sub_scope_id")
        year = request.form.get("year")
        sheet_name = request.form.get("sheet_name", "Fr-04.1")  # ค่าเริ่มต้นเป็น Fr-04.1

        # ตรวจสอบว่ามีไฟล์หรือไม่
        if "excel_file" not in request.files:
            response = make_response("")
            encoded_message = urllib.parse.quote("กรุณาเลือกไฟล์ Excel")
            trigger_data = {"showError": encoded_message, "closeModal": True}
            response.headers["HX-Trigger"] = json.dumps(trigger_data)
            return response

        file = request.files["excel_file"]

        # ตรวจสอบว่าเลือกไฟล์แล้ว
        if file.filename == "":
            response = make_response("")
            encoded_message = urllib.parse.quote("กรุณาเลือกไฟล์ Excel")
            trigger_data = {"showError": encoded_message, "closeModal": True}
            response.headers["HX-Trigger"] = json.dumps(trigger_data)
            return response

        # ตรวจสอบนามสกุลไฟล์
        if not file.filename.endswith((".xlsx", ".xls")):
            response = make_response("")
            encoded_message = urllib.parse.quote("กรุณาเลือกไฟล์ Excel (.xlsx หรือ .xls)")
            trigger_data = {"showError": encoded_message, "closeModal": True}
            response.headers["HX-Trigger"] = json.dumps(trigger_data)
            return response

        # กำหนด Scope ที่จะ import
        scopes_to_import = []

        if scope_id == "all":
            # Import ทุก Scope ที่ user มีสิทธิ์
            if current_user.ghg_scope_1:
                scope_1_list = Scope.objects(
                    ghg_scope=1,
                    ghg_sup_scope__in=current_user.ghg_scope_1,
                    campus=current_user.campus_id,
                    department=current_user.department_key,
                )
                scopes_to_import.extend(scope_1_list)

            if current_user.ghg_scope_2:
                scope_2_list = Scope.objects(
                    ghg_scope=2,
                    ghg_sup_scope__in=current_user.ghg_scope_2,
                    campus=current_user.campus_id,
                    department=current_user.department_key,
                )
                scopes_to_import.extend(scope_2_list)

            if current_user.ghg_scope_3:
                scope_3_list = Scope.objects(
                    ghg_scope=3,
                    ghg_sup_scope__in=current_user.ghg_scope_3,
                    campus=current_user.campus_id,
                    department=current_user.department_key,
                )
                scopes_to_import.extend(scope_3_list)
        elif sub_scope_id == "all":
            # Import ทุก Sub Scope ของ Scope ที่เลือก
            scope_num = int(scope_id)
            if scope_num == 1 and current_user.ghg_scope_1:
                scopes_to_import = Scope.objects(
                    ghg_scope=1,
                    ghg_sup_scope__in=current_user.ghg_scope_1,
                    campus=current_user.campus_id,
                    department=current_user.department_key,
                )
            elif scope_num == 2 and current_user.ghg_scope_2:
                scopes_to_import = Scope.objects(
                    ghg_scope=2,
                    ghg_sup_scope__in=current_user.ghg_scope_2,
                    campus=current_user.campus_id,
                    department=current_user.department_key,
                )
            elif scope_num == 3 and current_user.ghg_scope_3:
                scopes_to_import = Scope.objects(
                    ghg_scope=3,
                    ghg_sup_scope__in=current_user.ghg_scope_3,
                    campus=current_user.campus_id,
                    department=current_user.department_key,
                )
        else:
            # Import เฉพาะ Scope และ Sub Scope ที่เลือก
            scope = Scope.objects(
                ghg_scope=int(scope_id),
                ghg_sup_scope=int(sub_scope_id),
                campus=current_user.campus_id,
                department=current_user.department_key,
            ).first()

            if scope:
                scopes_to_import = [scope]

        if not scopes_to_import:
            response = make_response("")
            encoded_message = urllib.parse.quote("ไม่พบข้อมูล Scope ที่ระบุ")
            trigger_data = {"showError": encoded_message, "closeModal": True}
            response.headers["HX-Trigger"] = json.dumps(trigger_data)
            return response

        # อ่านไฟล์ Excel (data_only=True เพื่อดึงค่าแทนสูตร)
        wb = openpyxl.load_workbook(file, data_only=True)

        # ตรวจสอบว่ามี Sheet ที่ต้องการหรือไม่
        if sheet_name not in wb.sheetnames:
            response = make_response("")
            encoded_message = urllib.parse.quote(
                f"ไม่พบ Sheet ชื่อ '{sheet_name}' ในไฟล์ Excel"
            )
            trigger_data = {"showError": encoded_message, "closeModal": True}
            response.headers["HX-Trigger"] = json.dumps(trigger_data)
            return response

        ws = wb[sheet_name]

        # นับจำนวนแถวที่ประมวลผลสำเร็จ
        success_count = 0
        error_count = 0
        skipped_count = 0

        # วนลูปอ่านข้อมูลทีละแถว (เริ่มจากแถวที่ 2 สมมติว่าแถวแรกเป็นหัวตาราง)
        for row_idx, row in enumerate(
            ws.iter_rows(min_row=2, values_only=True), start=2
        ):
            try:
                # Column B = index 1 (0-based indexing)
                material_name = row[1] if len(row) > 1 else None

                # ข้ามแถวที่ Column B ว่างเปล่า
                if not material_name or str(material_name).strip() == "":
                    skipped_count += 1
                    continue

                # Column D = index 3 (0-based indexing)
                amount = row[3] if len(row) > 3 else None

                # ตรวจสอบว่ามีค่าหรือไม่
                if amount is None or amount == "":
                    skipped_count += 1
                    continue

                # แปลงเป็นตัวเลขถ้าเป็น string
                try:
                    amount = float(amount)
                except (ValueError, TypeError):
                    error_count += 1
                    continue

                # ค้นหา FormAndFormula ที่ตรงกับ material_name
                form_and_formula = FormAndFormula.objects(
                    material_name=str(material_name).strip()
                ).first()
                if not form_and_formula:
                    error_count += 1
                    continue

                # หารค่าด้วย 12 เพราะข้อมูลจาก Excel เป็นข้อมูลรายปี
                monthly_amount = amount / 12

                # บันทึกข้อมูลสำหรับแต่ละ Scope ที่เลือก
                for scope in scopes_to_import:
                    # ตรวจสอบว่า material นี้อยู่ใน head_table ของ scope หรือไม่
                    if str(material_name).strip() not in scope.head_table:
                        continue

                    # บันทึกข้อมูลสำหรับแต่ละเดือน (1-12)
                    for month_id in range(1, 13):
                        # ค้นหา InputType แรก (หรือที่ต้องการ) จาก form
                        if form_and_formula.input_types:
                            input_type = form_and_formula.input_types[0]
                            field = input_type.field

                            # ใช้ฟังก์ชันเดียวกันกับการกรอกมือจาก emissoins.py
                            material_data = {
                                "head": str(material_name).strip(),
                                "field": field,
                                "amount": str(monthly_amount),
                            }

                            save_material(
                                scope_id=scope.ghg_scope,
                                sub_scope_id=scope.ghg_sup_scope,
                                month_id=month_id,
                                year=int(year),
                                material_data=material_data,
                            )

                success_count += 1

            except Exception as e:
                print(f"Error processing row {row_idx}: {str(e)}")
                error_count += 1
                continue

        # ปิดไฟล์
        wb.close()

        # สร้างข้อความแจ้งผลลัพธ์
        message = f"นำเข้าข้อมูลสำเร็จ {success_count} รายการ"

        response = make_response("")
        encoded_message = urllib.parse.quote(message)
        trigger_data = {
            "showSuccess": encoded_message,
            "closeModal": True,
        }
        response.headers["HX-Trigger"] = json.dumps(trigger_data)
        response.headers["HX-Redirect"] = url_for("emissions_scope.emissions_scope")
        return response

    except Exception as e:
        response = make_response("")
        encoded_message = urllib.parse.quote(f"เกิดข้อผิดพลาด: {str(e)}")
        trigger_data = {"showError": encoded_message, "closeModal": True}
        response.headers["HX-Trigger"] = json.dumps(trigger_data)
        return response
