import os
import re
import sqlite3
import urllib.parse

from datetime import datetime
from io import BytesIO

from flask import (
    Flask,
    render_template,
    request,
    redirect,
    url_for,
    flash,
    jsonify,
    send_file
)

from flask_sqlalchemy import SQLAlchemy
from flask_login import (
    LoginManager,
    UserMixin,
    login_user,
    login_required,
    logout_user,
    current_user
)

from werkzeug.security import (
    generate_password_hash,
    check_password_hash
)

from openpyxl import Workbook
from openpyxl.styles import (
    Font,
    PatternFill,
    Alignment
)


# ============================================================
# APP SETUP
# ============================================================

app = Flask(__name__)

app.config["SECRET_KEY"] = os.environ.get(
    "SECRET_KEY",
    "eventpos-dev-secret-key-change-later"
)

app.config["SQLALCHEMY_DATABASE_URI"] = os.environ.get(
    "DATABASE_URL",
    "sqlite:///eventpos.db"
)
database_url = app.config["SQLALCHEMY_DATABASE_URI"]

if database_url.startswith("postgres://"):
    database_url = database_url.replace(
        "postgres://",
        "postgresql://",
        1
    )

app.config["SQLALCHEMY_DATABASE_URI"] = database_url

app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False


db = SQLAlchemy(app)

login_manager = LoginManager(app)
login_manager.login_view = "login"
login_manager.login_message = "Please login to continue."


# ============================================================
# DATABASE MODELS
# ============================================================

class User(UserMixin, db.Model):

    id = db.Column(
        db.Integer,
        primary_key=True
    )

    username = db.Column(
        db.String(100),
        unique=True,
        nullable=False
    )

    password_hash = db.Column(
        db.String(255),
        nullable=False
    )

    role = db.Column(
        db.String(30),
        nullable=False,
        default="cashier"
    )

    is_active_user = db.Column(
        db.Boolean,
        nullable=False,
        default=True
    )


class Product(db.Model):

    id = db.Column(
        db.Integer,
        primary_key=True
    )

    name = db.Column(
        db.String(200),
        nullable=False
    )

    price = db.Column(
        db.Float,
        nullable=False
    )

    discount_percent = db.Column(
        db.Float,
        nullable=False,
        default=0
    )

    is_active = db.Column(
        db.Boolean,
        nullable=False,
        default=True
    )


class Customer(db.Model):

    id = db.Column(
        db.Integer,
        primary_key=True
    )

    mobile_number = db.Column(
        db.String(30),
        unique=True,
        nullable=False
    )

    customer_name = db.Column(
        db.String(200),
        nullable=False
    )

    created_date = db.Column(
        db.DateTime,
        nullable=False,
        default=datetime.now
    )

    updated_date = db.Column(
        db.DateTime,
        nullable=False,
        default=datetime.now,
        onupdate=datetime.now
    )


class Sale(db.Model):

    id = db.Column(
        db.Integer,
        primary_key=True
    )

    sale_date = db.Column(
        db.DateTime,
        nullable=False,
        default=datetime.now
    )

    customer_name = db.Column(
        db.String(200),
        nullable=True
    )

    customer_number = db.Column(
        db.String(30),
        nullable=True
    )

    payment_mode = db.Column(
        db.String(20),
        nullable=False
    )

    total_amount = db.Column(
        db.Float,
        nullable=False
    )

    items = db.relationship(
        "SaleItem",
        backref="sale",
        lazy=True,
        cascade="all, delete-orphan"
    )


class SaleItem(db.Model):

    id = db.Column(
        db.Integer,
        primary_key=True
    )

    sale_id = db.Column(
        db.Integer,
        db.ForeignKey("sale.id"),
        nullable=False
    )

    product_id = db.Column(
        db.Integer,
        nullable=True
    )

    product_name = db.Column(
        db.String(200),
        nullable=False
    )

    quantity = db.Column(
        db.Integer,
        nullable=False
    )

    unit_price = db.Column(
        db.Float,
        nullable=False
    )

    total_price = db.Column(
        db.Float,
        nullable=False
    )


# ============================================================
# LOGIN
# ============================================================

@login_manager.user_loader
def load_user(user_id):

    return db.session.get(
        User,
        int(user_id)
    )


def create_default_admin():

    existing = User.query.filter_by(
        username="admin"
    ).first()

    if existing:
        return

    admin = User(
        username="admin",
        password_hash=generate_password_hash(
            "admin123"
        ),
        role="admin"
    )

    db.session.add(admin)
    db.session.commit()


# ============================================================
# HELPERS
# ============================================================

def calculate_final_price(
    price,
    discount_percent
):

    discount_percent = (
        discount_percent or 0
    )

    final_price = (
        price
        - (
            price
            * discount_percent
            / 100
        )
    )

    return round(
        final_price,
        2
    )


def normalize_whatsapp_number(number):

    if not number:
        return ""

    digits = re.sub(
        r"\D",
        "",
        number
    )

    if len(digits) == 10:
        digits = "91" + digits

    elif (
        len(digits) == 11
        and digits.startswith("0")
    ):
        digits = "91" + digits[1:]

    return digits


def build_whatsapp_message(
    sale
):

    lines = []

    lines.append(
        "Thank you for your order! 😊"
    )

    lines.append("")

    lines.append(
        f"Order No: #{sale.id}"
    )

    lines.append(
        "Date: "
        + sale.sale_date.strftime(
            "%d-%m-%Y %I:%M %p"
        )
    )

    if (
        sale.customer_name
        and sale.customer_name != "Walk-in Customer"
    ):

        lines.append(
            f"Customer: {sale.customer_name}"
        )

    lines.append("")
    lines.append("Order Details:")
    lines.append("----------------------")

    for item in sale.items:

        lines.append(
            f"{item.product_name} x {item.quantity}"
        )

        lines.append(
            f"₹{item.unit_price:.2f} x {item.quantity}"
            f" = ₹{item.total_price:.2f}"
        )

    lines.append("----------------------")

    lines.append(
        f"Total: ₹{sale.total_amount:.2f}"
    )

    lines.append(
        f"Payment: {sale.payment_mode}"
    )

    lines.append("")

    lines.append(
        "Thank you for shopping with us!"
    )

    return "\n".join(lines)


def admin_required():

    return (
        current_user.is_authenticated
        and current_user.role == "admin"
    )


# ============================================================
# LOGIN ROUTES
# ============================================================

@app.route(
    "/login",
    methods=[
        "GET",
        "POST"
    ]
)
def login():

    if current_user.is_authenticated:

        return redirect(
            url_for("dashboard")
        )

    if request.method == "POST":

        username = (
            request.form
            .get("username", "")
            .strip()
        )

        password = (
            request.form
            .get("password", "")
        )

        user = User.query.filter_by(
            username=username
        ).first()

        if (
            user
            and user.is_active_user
            and check_password_hash(
                user.password_hash,
                password
            )
        ):

            login_user(user)

            return redirect(
                url_for("dashboard")
            )

        flash(
            "Invalid username or password.",
            "danger"
        )

    return render_template(
        "login.html"
    )


@app.route("/logout")
@login_required
def logout():

    logout_user()

    return redirect(
        url_for("login")
    )


# ============================================================
# DASHBOARD
# ============================================================

@app.route("/")
@login_required
def dashboard():

    today = datetime.now().date()

    sales = Sale.query.all()

    today_sales = [
        sale
        for sale in sales
        if sale.sale_date.date() == today
    ]

    orders = len(today_sales)

    cash_total = sum(
        sale.total_amount
        for sale in today_sales
        if sale.payment_mode == "Cash"
    )

    upi_total = sum(
        sale.total_amount
        for sale in today_sales
        if sale.payment_mode == "UPI"
    )

    total_sales = sum(
        sale.total_amount
        for sale in today_sales
    )

    top_items = {}

    for sale in today_sales:

        for item in sale.items:

            top_items[
                item.product_name
            ] = (
                top_items.get(
                    item.product_name,
                    0
                )
                + item.quantity
            )

    top_items = sorted(
        top_items.items(),
        key=lambda x: x[1],
        reverse=True
    )[:5]

    return render_template(
        "dashboard.html",
        orders=orders,
        cash_total=cash_total,
        upi_total=upi_total,
        total_sales=total_sales,
        top_items=top_items
    )


# ============================================================
# CUSTOMER LOOKUP
# ============================================================

@app.route(
    "/api/customer/<mobile>"
)
@login_required
def customer_lookup(mobile):

    mobile = mobile.strip()

    customer = Customer.query.filter_by(
        mobile_number=mobile
    ).first()

    if customer:

        return jsonify({
            "found": True,
            "name": customer.customer_name
        })

    old_sale = (
        Sale.query
        .filter(
            Sale.customer_number == mobile,
            Sale.customer_name.isnot(None),
            Sale.customer_name != "",
            Sale.customer_name != "Walk-in Customer"
        )
        .order_by(
            Sale.id.desc()
        )
        .first()
    )

    if old_sale:

        return jsonify({
            "found": True,
            "name": old_sale.customer_name
        })

    return jsonify({
        "found": False,
        "name": ""
    })


# ============================================================
# NEW SALE
# ============================================================

@app.route("/sale")
@login_required
def new_sale():

    products = (
        Product.query
        .filter_by(
            is_active=True
        )
        .order_by(
            Product.name
        )
        .all()
    )

    product_data = []

    for product in products:

        final_price = calculate_final_price(
            product.price,
            product.discount_percent
        )

        product_data.append({
            "id": product.id,
            "name": product.name,
            "price": product.price,
            "discount": product.discount_percent,
            "final_price": final_price
        })

    return render_template(
        "sale.html",
        products=product_data
    )


@app.route(
    "/api/complete-sale",
    methods=["POST"]
)
@login_required
def complete_sale():

    data = request.get_json(
        silent=True
    )

    if not data:

        return jsonify({
            "success": False,
            "message": "Invalid request."
        }), 400

    customer_number = (
        data.get(
            "customer_number",
            ""
        )
        .strip()
    )

    customer_name = (
        data.get(
            "customer_name",
            ""
        )
        .strip()
    )

    payment_mode = (
        data.get(
            "payment_mode",
            ""
        )
        .strip()
    )

    items = data.get(
        "items",
        []
    )

    if not items:

        return jsonify({
            "success": False,
            "message": "Please select at least one item."
        }), 400

    if payment_mode not in [
        "Cash",
        "UPI"
    ]:

        return jsonify({
            "success": False,
            "message": "Please select Cash or UPI."
        }), 400

    if (
        customer_number
        and not customer_name
    ):

        return jsonify({
            "success": False,
            "message": "Please enter customer name."
        }), 400

    if not customer_name:

        customer_name = (
            "Walk-in Customer"
        )

    total_amount = 0

    prepared_items = []

    for item in items:

        product_id = item.get("product_id")

        quantity = int(
            item.get(
                "quantity",
                0
            )
        )

        if quantity <= 0:
            continue

        product = db.session.get(
            Product,
            product_id
        )

        if not product:
            continue

        if not product.is_active:
            continue

        final_price = (
            calculate_final_price(
                product.price,
                product.discount_percent
            )
        )

        item_total = round(
            final_price
            * quantity,
            2
        )

        total_amount += (
            item_total
        )

        prepared_items.append({
            "product_id": product.id,
            "product_name": product.name,
            "quantity": quantity,
            "unit_price": final_price,
            "total_price": item_total
        })

    if not prepared_items:

        return jsonify({
            "success": False,
            "message": "No valid products found."
        }), 400

    sale = Sale(
        sale_date=datetime.now(),
        customer_name=customer_name,
        customer_number=customer_number,
        payment_mode=payment_mode,
        total_amount=round(
            total_amount,
            2
        )
    )

    db.session.add(sale)
    db.session.flush()

    for item in prepared_items:

        sale_item = SaleItem(
            sale_id=sale.id,
            product_id=item[
                "product_id"
            ],
            product_name=item[
                "product_name"
            ],
            quantity=item[
                "quantity"
            ],
            unit_price=item[
                "unit_price"
            ],
            total_price=item[
                "total_price"
            ]
        )

        db.session.add(
            sale_item
        )

    if (
        customer_number
        and customer_name
        and customer_name != "Walk-in Customer"
    ):

        customer = Customer.query.filter_by(
            mobile_number=customer_number
        ).first()

        if customer:

            customer.customer_name = (
                customer_name
            )

            customer.updated_date = (
                datetime.now()
            )

        else:

            customer = Customer(
                mobile_number=customer_number,
                customer_name=customer_name,
                created_date=datetime.now(),
                updated_date=datetime.now()
            )

            db.session.add(
                customer
            )

    db.session.commit()

    whatsapp_url = ""

    if customer_number:

        whatsapp_number = (
            normalize_whatsapp_number(
                customer_number
            )
        )

        if whatsapp_number:

            whatsapp_message = (
                build_whatsapp_message(
                    sale
                )
            )

            whatsapp_url = (
                "https://wa.me/"
                + whatsapp_number
                + "?text="
                + urllib.parse.quote(
                    whatsapp_message
                )
            )

    return jsonify({
        "success": True,
        "sale_id": sale.id,
        "total": sale.total_amount,
        "whatsapp_url": whatsapp_url
    })


# ============================================================
# PRODUCTS
# ============================================================

@app.route(
    "/products",
    methods=[
        "GET",
        "POST"
    ]
)
@login_required
def products():

    if not admin_required():

        flash(
            "Only admin can manage products.",
            "danger"
        )

        return redirect(
            url_for("new_sale")
        )

    if request.method == "POST":

        name = (
            request.form
            .get("name", "")
            .strip()
        )

        price_text = (
            request.form
            .get("price", "")
            .strip()
        )

        discount_text = (
            request.form
            .get(
                "discount_percent",
                "0"
            )
            .strip()
        )

        if not name:

            flash(
                "Product name is required.",
                "danger"
            )

            return redirect(
                url_for("products")
            )

        try:

            price = float(
                price_text
            )

            discount = float(
                discount_text or 0
            )

        except ValueError:

            flash(
                "Price and discount must be valid numbers.",
                "danger"
            )

            return redirect(
                url_for("products")
            )

        if price <= 0:

            flash(
                "Price must be greater than zero.",
                "danger"
            )

            return redirect(
                url_for("products")
            )

        if (
            discount < 0
            or discount > 100
        ):

            flash(
                "Discount must be between 0 and 100.",
                "danger"
            )

            return redirect(
                url_for("products")
            )

        product = Product(
            name=name,
            price=price,
            discount_percent=discount,
            is_active=True
        )

        db.session.add(product)
        db.session.commit()

        flash(
            "Product saved successfully.",
            "success"
        )

        return redirect(
            url_for("products")
        )

    products_list = (
        Product.query
        .order_by(
            Product.name
        )
        .all()
    )

    return render_template(
        "products.html",
        products=products_list,
        calculate_final_price=calculate_final_price
    )


@app.route(
    "/products/<int:product_id>/edit",
    methods=["POST"]
)
@login_required
def edit_product(product_id):

    if not admin_required():

        return redirect(
            url_for("new_sale")
        )

    product = db.session.get(
        Product,
        product_id
    )

    if not product:

        flash(
            "Product not found.",
            "danger"
        )

        return redirect(
            url_for("products")
        )

    name = (
        request.form
        .get("name", "")
        .strip()
    )

    try:

        price = float(
            request.form
            .get("price", "0")
        )

        discount = float(
            request.form
            .get(
                "discount_percent",
                "0"
            )
        )

    except ValueError:

        flash(
            "Invalid price or discount.",
            "danger"
        )

        return redirect(
            url_for("products")
        )

    if not name:

        flash(
            "Product name is required.",
            "danger"
        )

        return redirect(
            url_for("products")
        )

    if price <= 0:

        flash(
            "Price must be greater than zero.",
            "danger"
        )

        return redirect(
            url_for("products")
        )

    if (
        discount < 0
        or discount > 100
    ):

        flash(
            "Discount must be between 0 and 100.",
            "danger"
        )

        return redirect(
            url_for("products")
        )

    product.name = name
    product.price = price
    product.discount_percent = discount

    db.session.commit()

    flash(
        "Product updated successfully.",
        "success"
    )

    return redirect(
        url_for("products")
    )


@app.route(
    "/products/<int:product_id>/toggle",
    methods=["POST"]
)
@login_required
def toggle_product(product_id):

    if not admin_required():

        return redirect(
            url_for("new_sale")
        )

    product = db.session.get(
        Product,
        product_id
    )

    if not product:

        flash(
            "Product not found.",
            "danger"
        )

        return redirect(
            url_for("products")
        )

    product.is_active = (
        not product.is_active
    )

    db.session.commit()

    return redirect(
        url_for("products")
    )


# ============================================================
# SALES HISTORY
# ============================================================

@app.route("/sales-history")
@login_required
def sales_history():

    search_query = request.args.get("search", "").strip()
    payment_filter = request.args.get("payment", "").strip()
    date_filter = request.args.get("date", "").strip()

    sales = Sale.query.order_by(
        Sale.id.desc()
    ).all()

    # Search by Order ID, Customer Name or Mobile Number
    if search_query:

        search_lower = search_query.lower()

        filtered_sales = []

        for sale in sales:

            order_match = str(sale.id) == search_query

            customer_match = (
                search_lower
                in (sale.customer_name or "").lower()
            )

            mobile_match = (
                search_query
                in (sale.customer_number or "")
            )

            if (
                order_match
                or customer_match
                or mobile_match
            ):
                filtered_sales.append(sale)

        sales = filtered_sales


    # Filter by Cash / UPI
    if payment_filter:

        sales = [
            sale
            for sale in sales
            if (sale.payment_mode or "").lower()
            == payment_filter.lower()
        ]


    # Filter by Date
    if date_filter:

        sales = [
            sale
            for sale in sales
            if sale.sale_date
            and sale.sale_date.strftime("%Y-%m-%d")
            == date_filter
        ]


    # Calculate totals after filtering
    cash_total = sum(
        float(sale.total_amount or 0)
        for sale in sales
        if (sale.payment_mode or "").lower() == "cash"
    )

    upi_total = sum(
        float(sale.total_amount or 0)
        for sale in sales
        if (sale.payment_mode or "").lower() == "upi"
    )

    grand_total = sum(
        float(sale.total_amount or 0)
        for sale in sales
    )


    return render_template(
        "sales_history.html",
        sales=sales,
        cash_total=cash_total,
        upi_total=upi_total,
        grand_total=grand_total,
        search_query=search_query,
        payment_filter=payment_filter,
        date_filter=date_filter
    )


# ============================================================
# EXCEL EXPORT
# ============================================================

@app.route("/export-sales")
@login_required
def export_sales():

    search_query = request.args.get("search", "").strip()
    payment_filter = request.args.get("payment", "").strip()
    date_filter = request.args.get("date", "").strip()

    sales = Sale.query.order_by(
        Sale.id.desc()
    ).all()

    # --------------------------------------------------------
    # SEARCH FILTER
    # --------------------------------------------------------

    if search_query:

        search_lower = search_query.lower()

        filtered_sales = []

        for sale in sales:

            order_match = (
                str(sale.id) == search_query
            )

            customer_match = (
                search_lower
                in (sale.customer_name or "").lower()
            )

            mobile_match = (
                search_query
                in (sale.customer_number or "")
            )

            if (
                order_match
                or customer_match
                or mobile_match
            ):
                filtered_sales.append(sale)

        sales = filtered_sales


    # --------------------------------------------------------
    # PAYMENT FILTER
    # --------------------------------------------------------

    if payment_filter:

        sales = [
            sale
            for sale in sales
            if (sale.payment_mode or "").lower()
            == payment_filter.lower()
        ]


    # --------------------------------------------------------
    # DATE FILTER
    # --------------------------------------------------------

    if date_filter:

        sales = [
            sale
            for sale in sales
            if (
                sale.sale_date
                and sale.sale_date.strftime("%Y-%m-%d")
                == date_filter
            )
        ]


    # --------------------------------------------------------
    # CREATE EXCEL
    # --------------------------------------------------------

    workbook = Workbook()

    sheet = workbook.active
    sheet.title = "Sales Report"


    headers = [
        "Order ID",
        "Date",
        "Customer Name",
        "Mobile Number",
        "Payment Mode",
        "Product",
        "Quantity",
        "Unit Price",
        "Item Total",
        "Order Total"
    ]

    sheet.append(headers)


    # --------------------------------------------------------
    # HEADER STYLE
    # --------------------------------------------------------

    for cell in sheet[1]:

        cell.font = Font(
            bold=True
        )

        cell.alignment = Alignment(
            horizontal="center"
        )


    # --------------------------------------------------------
    # ADD SALES
    # --------------------------------------------------------

    for sale in sales:

        for item in sale.items:

            sheet.append([
                sale.id,

                sale.sale_date.strftime(
                    "%d-%m-%Y %I:%M %p"
                ),

                sale.customer_name
                or "Walk-in Customer",

                sale.customer_number
                or "",

                sale.payment_mode,

                item.product_name,

                item.quantity,

                round(
                    float(item.unit_price or 0),
                    2
                ),

                round(
                    float(item.total_price or 0),
                    2
                ),

                round(
                    float(sale.total_amount or 0),
                    2
                )
            ])


    # --------------------------------------------------------
    # COLUMN WIDTHS
    # --------------------------------------------------------

    widths = {
        "A": 12,
        "B": 22,
        "C": 25,
        "D": 18,
        "E": 15,
        "F": 30,
        "G": 12,
        "H": 15,
        "I": 15,
        "J": 15
    }

    for column, width in widths.items():

        sheet.column_dimensions[
            column
        ].width = width


    # --------------------------------------------------------
    # CREATE FILE IN MEMORY
    # --------------------------------------------------------

    output = BytesIO()

    workbook.save(output)

    output.seek(0)


    filename = (
        "Sales_Report_"
        + datetime.now().strftime(
            "%d-%m-%Y_%H-%M-%S"
        )
        + ".xlsx"
    )


    # --------------------------------------------------------
    # DOWNLOAD
    # --------------------------------------------------------

    return send_file(
        output,
        as_attachment=True,
        download_name=filename,
        mimetype=(
            "application/vnd.openxmlformats-"
            "officedocument.spreadsheetml.sheet"
        )
    )
# ============================================================
# INITIALIZE DATABASE
# ============================================================

with app.app_context():

    db.create_all()

    create_default_admin()


# ============================================================
# RUN
# ============================================================

if __name__ == "__main__":

    app.run(
        host="127.0.0.1",
        port=5000,
        debug=True
    )