// ==========================================================================
// 1. CONFIGURATION ET INITIALISATION DE FIREBASE
// ==========================================================================

// Remplace ces valeurs par celles de ton projet Firebase.
const firebaseConfig = {
  apiKey: "AIzaSyALh3_J6f48D2269aH6tQGqrP6s08AFKAI",
  authDomain: "gestion-boutique-98d0f.firebaseapp.com",
  projectId: "gestion-boutique-98d0f",
  storageBucket: "gestion-boutique-98d0f.firebasestorage.app",
  messagingSenderId: "647096731225",
  appId: "1:647096731225:web:2c5002e47e7e7d09f99348"
};

if (typeof firebase === "undefined") {
    throw new Error(
        "Firebase n'est pas chargé. Vérifie les scripts Firebase dans ton HTML."
    );
}

if (typeof XLSX === "undefined") {
    console.warn(
        "SheetJS n'est pas chargé. L'exportation Excel sera indisponible."
    );
}

if (!firebase.apps.length) {
    firebase.initializeApp(firebaseConfig);
}

const db = firebase.firestore();


// ==========================================================================
// 2. DICTIONNAIRE DE TRADUCTION FR / EN
// ==========================================================================

const i18n = {
    FR: {
        appTitle: "Gestionnaire de Commandes Pro",
        newOrderTitle: "Nouvelle Commande",
        editOrderTitle: "Modifier la Commande",

        lblOrderNum: "Numéro de Commande",
        lblOrderDate: "Date de Commande",
        lblClientName: "Nom du Client",
        lblClientContact: "Contact Client",
        lblDeliveryLoc: "Lieu de Livraison",
        lblPayStatus: "Statut Paiement",
        lblDelivStatus: "Statut Préparation & Livraison",

        optUnpaid: "Non payé",
        optPartial: "Payé partiellement",
        optPaid: "Payé",

        optNew: "Nouvelle commande",
        optPreparing: "En cours de préparation",
        optShipping: "En cours de livraison",
        optDelivered: "Livré",

        titleProducts: "Produits de la commande",
        btnAddProduct: "+ Ajouter un produit",

        lblSubtotal: "Sous-total ($)",
        lblDelivFee: "Frais de livraison ($)",
        lblTotal: "Prix Total ($)",

        btnAddCategory: "Ajouter d'autres catégories",
        btnSaveOrder: "Enregistrer la Commande",
        btnCancel: "Annuler",

        titleReports: "Tableau des Rapports & Performances",
        statTotalOrders: "Nombre Total de Commandes",
        statTotalRevenue: "Chiffre d'Affaires Cumulé",
        statPendingDeliv: "Livraisons en Cours",

        titleDailyReport: "Rapport Journalier Automatique",
        noDataToday: "Aucune commande enregistrée pour aujourd'hui.",

        titleOrdersTable:
            "Registre Global des Commandes (Mode Grille Excel)",

        btnExportExcel: "Exporter les Données en Excel (.xlsx)",

        titleFilters: "Filtres Avancés",
        filterAllPay: "Tous les statuts de paiement",
        filterAllDeliv: "Tous les statuts de livraison",

        placeSearch: "Rechercher (Client, Commande, Produit...)",

        thOrderNum: "N° Commande",
        thDate: "Date",
        thClient: "Client & Contact",
        thLocation: "Lieu Liv.",
        thProducts: "Produits (Nom x Qté)",
        thSubtotal: "Sous-Total",
        thDelivFee: "Frais Liv.",
        thTotal: "Prix Total",
        thPayStatus: "Statut Paiement",
        thDelivStatus: "Statut Logistique",
        thActions: "Actions",

        lblProdName: "Nom du produit",
        lblProdPrice: "Prix Unit.",
        lblProdQty: "Qté",
        lblProdSub: "Total",

        msgSaved: "Commande enregistrée avec succès !",
        msgDeleted: "Commande supprimée !",
        msgConfirmDelete: "Supprimer cette commande ?",
        msgNoOrders: "Aucune commande à exporter.",
        msgExportError: "Impossible de créer le fichier Excel.",
        msgLoadError: "Impossible de charger les commandes.",
        msgSaveError: "Erreur lors de l'enregistrement.",
        msgDeleteError: "Erreur lors de la suppression.",
        msgCategory:
            "La fonctionnalité d'ajout de catégories n'est pas encore implémentée.",

        phraseReport:
            "Aujourd'hui, {count} commande(s) ont été traitées pour un montant total de {total} $. {pending} livraison(s) sont actuellement en attente."
    },

    EN: {
        appTitle: "Pro Order Manager",
        newOrderTitle: "New Order",
        editOrderTitle: "Edit Order",

        lblOrderNum: "Order Number",
        lblOrderDate: "Order Date",
        lblClientName: "Client Name",
        lblClientContact: "Client Contact",
        lblDeliveryLoc: "Delivery Location",
        lblPayStatus: "Payment Status",
        lblDelivStatus: "Preparation & Delivery Status",

        optUnpaid: "Unpaid",
        optPartial: "Partially Paid",
        optPaid: "Paid",

        optNew: "New Order",
        optPreparing: "In Preparation",
        optShipping: "In Shipping",
        optDelivered: "Delivered",

        titleProducts: "Order Products",
        btnAddProduct: "+ Add Product",

        lblSubtotal: "Subtotal ($)",
        lblDelivFee: "Delivery Fee ($)",
        lblTotal: "Total Price ($)",

        btnAddCategory: "Add Other Categories",
        btnSaveOrder: "Save Order",
        btnCancel: "Cancel",

        titleReports: "Reports & Performance Dashboard",
        statTotalOrders: "Total Number of Orders",
        statTotalRevenue: "Cumulative Revenue",
        statPendingDeliv: "Deliveries in Progress",

        titleDailyReport: "Automatic Daily Report",
        noDataToday: "No orders registered for today.",

        titleOrdersTable: "Global Order Register (Excel Grid Mode)",

        btnExportExcel: "Export Data to Excel (.xlsx)",

        titleFilters: "Advanced Filters",
        filterAllPay: "All payment statuses",
        filterAllDeliv: "All delivery statuses",

        placeSearch: "Search (Client, Order, Product...)",

        thOrderNum: "Order No.",
        thDate: "Date",
        thClient: "Client & Contact",
        thLocation: "Delivery Place",
        thProducts: "Products (Name x Qty)",
        thSubtotal: "Subtotal",
        thDelivFee: "Delivery Fee",
        thTotal: "Total Price",
        thPayStatus: "Payment Status",
        thDelivStatus: "Logistics Status",
        thActions: "Actions",

        lblProdName: "Product Name",
        lblProdPrice: "Unit Price",
        lblProdQty: "Qty",
        lblProdSub: "Total",

        msgSaved: "Order successfully saved!",
        msgDeleted: "Order deleted!",
        msgConfirmDelete: "Delete this order?",
        msgNoOrders: "No orders to export.",
        msgExportError: "Unable to create the Excel file.",
        msgLoadError: "Unable to load orders.",
        msgSaveError: "Error while saving the order.",
        msgDeleteError: "Error while deleting the order.",
        msgCategory:
            "The category extension feature has not been implemented yet.",

        phraseReport:
            "Today, {count} order(s) were processed for a total amount of {total} $. {pending} delivery(ies) are currently pending."
    }
};

let currentLang = "FR";


// ==========================================================================
// 3. FONCTIONS UTILITAIRES
// ==========================================================================

// Convertir une valeur en nombre sans produire NaN.
function toNumber(value) {
    const number = Number(value);
    return Number.isFinite(number) ? number : 0;
}

// Échapper les caractères HTML pour éviter l'injection de code.
function escapeHTML(value) {
    return String(value ?? "").replace(/[&<>"']/g, char => {
        const entities = {
            "&": "&amp;",
            "<": "&lt;",
            ">": "&gt;",
            '"': "&quot;",
            "'": "&#39;"
        };

        return entities[char];
    });
}

// Obtenir la date locale au format YYYY-MM-DD.
function getLocalDate() {
    const now = new Date();

    const year = now.getFullYear();
    const month = String(now.getMonth() + 1).padStart(2, "0");
    const day = String(now.getDate()).padStart(2, "0");

    return `${year}-${month}-${day}`;
}

// Formater les montants.
function formatMoney(value) {
    return toNumber(value).toFixed(2);
}

// Lire une traduction.
function t(key) {
    return i18n[currentLang][key] ?? key;
}

// Vérifier qu'une commande possède une liste de produits valide.
function normalizeProducts(products) {
    if (!Array.isArray(products)) {
        return [];
    }

    return products.map(product => ({
        name: String(product?.name ?? ""),
        price: toNumber(product?.price),
        qty: Math.max(0, toNumber(product?.qty))
    }));
}


// ==========================================================================
// 4. TRADUCTION DE L'INTERFACE
// ==========================================================================

function switchLanguage(lang) {
    if (!i18n[lang]) {
        return;
    }

    currentLang = lang;

    document.querySelectorAll("#btn-fr, #btn-en").forEach(button => {
        button.classList.toggle(
            "active",
            button.id === `btn-${lang.toLowerCase()}`
        );
    });

    // Traduire les éléments statiques.
    document.querySelectorAll("[data-i18n]").forEach(element => {
        const key = element.getAttribute("data-i18n");

        if (i18n[currentLang][key] !== undefined) {
            element.textContent = i18n[currentLang][key];
        }
    });

    // Traduire les placeholders.
    document.querySelectorAll("[data-i18n-holder]").forEach(element => {
        const key = element.getAttribute("data-i18n-holder");

        if (i18n[currentLang][key] !== undefined) {
            element.setAttribute(
                "placeholder",
                i18n[currentLang][key]
            );
        }
    });

    renderTable();
    updateReportsAndPhrases();
}


// ==========================================================================
// 5. VARIABLES PRINCIPALES
// ==========================================================================

let allOrders = [];
let unsubscribeOrders = null;

const productsContainer = document.getElementById("products-container");
const tbody = document.getElementById("orders-tbody");


// ==========================================================================
// 6. GESTION DES PRODUITS
// ==========================================================================

function createProductRow(productData = {}) {
    const product = {
        name: String(productData.name ?? ""),
        price: toNumber(productData.price),
        qty: Math.max(1, toNumber(productData.qty) || 1)
    };

    const row = document.createElement("div");
    row.className = "product-row";

    row.innerHTML = `
        <div class="form-group">
            <label data-i18n="lblProdName">
                ${escapeHTML(t("lblProdName"))}
            </label>

            <input
                type="text"
                class="prod-name"
                value="${escapeHTML(product.name)}"
                required
            >
        </div>

        <div class="form-group">
            <label data-i18n="lblProdPrice">
                ${escapeHTML(t("lblProdPrice"))}
            </label>

            <input
                type="number"
                class="prod-price"
                min="0"
                step="0.01"
                value="${product.price}"
                required
            >
        </div>

        <div class="form-group">
            <label data-i18n="lblProdQty">
                ${escapeHTML(t("lblProdQty"))}
            </label>

            <input
                type="number"
                class="prod-qty"
                min="1"
                step="1"
                value="${product.qty}"
                required
            >
        </div>

        <div class="form-group">
            <label data-i18n="lblProdSub">
                ${escapeHTML(t("lblProdSub"))}
            </label>

            <input
                type="number"
                class="prod-row-subtotal"
                value="${formatMoney(product.price * product.qty)}"
                readonly
            >
        </div>

        <button
            type="button"
            class="btn btn-danger btn-sm btn-remove-prod"
            aria-label="Remove product"
        >
            ✕
        </button>
    `;

    const priceInput = row.querySelector(".prod-price");
    const qtyInput = row.querySelector(".prod-qty");
    const subtotalInput = row.querySelector(".prod-row-subtotal");

    function calculateRowTotal() {
        const price = Math.max(0, toNumber(priceInput.value));
        const qty = Math.max(0, toNumber(qtyInput.value));

        subtotalInput.value = formatMoney(price * qty);

        calculateOrderTotals();
    }

    priceInput.addEventListener("input", calculateRowTotal);
    qtyInput.addEventListener("input", calculateRowTotal);

    row.querySelector(".btn-remove-prod").addEventListener(
        "click",
        () => {
            row.remove();
            calculateOrderTotals();
        }
    );

    productsContainer.appendChild(row);

    return row;
}


// ==========================================================================
// 7. CALCUL DES MONTANTS
// ==========================================================================

function calculateOrderTotals() {
    let subtotal = 0;

    document.querySelectorAll(".product-row").forEach(row => {
        const price = Math.max(
            0,
            toNumber(row.querySelector(".prod-price")?.value)
        );

        const qty = Math.max(
            0,
            toNumber(row.querySelector(".prod-qty")?.value)
        );

        const lineTotal = price * qty;

        const subtotalInput = row.querySelector(".prod-row-subtotal");

        if (subtotalInput) {
            subtotalInput.value = formatMoney(lineTotal);
        }

        subtotal += lineTotal;
    });

    const subtotalElement = document.getElementById("order-subtotal");
    const deliveryFeeElement = document.getElementById("delivery-fee");
    const totalElement = document.getElementById("order-total");

    if (!subtotalElement || !deliveryFeeElement || !totalElement) {
        return;
    }

    const deliveryFee = Math.max(
        0,
        toNumber(deliveryFeeElement.value)
    );

    subtotalElement.value = formatMoney(subtotal);
    totalElement.value = formatMoney(subtotal + deliveryFee);
}


// ==========================================================================
// 8. NUMÉRO DE COMMANDE
// ==========================================================================

function generateOrderNumber() {
    const maxNumber = allOrders.reduce((max, order) => {
        const match = String(order.number || "").match(/^CMD-(\d+)$/);

        if (match) {
            return Math.max(max, Number(match[1]));
        }

        return max;
    }, 0);

    return `CMD-${maxNumber + 1}`;
}

// ==========================================================================
// 9. RÉINITIALISATION DU FORMULAIRE
// ==========================================================================

function initForm() {
    const form = document.getElementById("order-form");

    form.reset();

    document.getElementById("order-id").value = "";

    document.getElementById("order-number").value =
        generateOrderNumber();

    document.getElementById("order-date").value = getLocalDate();

    productsContainer.innerHTML = "";

    createProductRow();

    document.getElementById("btn-cancel-edit")
        .classList.add("hidden");

    const title = document.getElementById("form-title");

    title.setAttribute("data-i18n", "newOrderTitle");
    title.textContent = t("newOrderTitle");

    calculateOrderTotals();
}


// ==========================================================================
// 10. ÉVÉNEMENTS DU FORMULAIRE
// ==========================================================================

document.getElementById("btn-add-product").addEventListener(
    "click",
    () => createProductRow()
);

document.getElementById("delivery-fee").addEventListener(
    "input",
    calculateOrderTotals
);

document.getElementById("btn-add-category").addEventListener(
    "click",
    () => {
        alert(t("msgCategory"));
    }
);


// ==========================================================================
// 11. RÉCUPÉRATION DES PRODUITS SAISIS
// ==========================================================================

function getProductsFromForm() {
    const rows = Array.from(
        document.querySelectorAll(".product-row")
    );

    return rows.map(row => ({
        name: row.querySelector(".prod-name").value.trim(),
        price: Math.max(
            0,
            toNumber(row.querySelector(".prod-price").value)
        ),
        qty: Math.max(
            1,
            Math.floor(toNumber(row.querySelector(".prod-qty").value))
        )
    }));
}


// ==========================================================================
// 12. ENREGISTREMENT ET MODIFICATION DES COMMANDES
// ==========================================================================

document.getElementById("order-form").addEventListener(
    "submit",
    async event => {
        event.preventDefault();

        const orderId = document.getElementById("order-id").value;

        const products = getProductsFromForm();

        if (products.length === 0) {
            alert(
                currentLang === "FR"
                    ? "Ajoute au moins un produit."
                    : "Add at least one product."
            );

            return;
        }

        if (products.some(product => !product.name.trim())) {
            alert(
                currentLang === "FR"
                    ? "Tous les produits doivent avoir un nom."
                    : "Every product must have a name."
            );

            return;
        }

        const clientName = document.getElementById("client-name")
            .value.trim();

        if (!clientName) {
            alert(
                currentLang === "FR"
                    ? "Le nom du client est obligatoire."
                    : "The client name is required."
            );

            return;
        }

        const subtotal = products.reduce(
            (sum, product) => sum + product.price * product.qty,
            0
        );

        const deliveryFee = Math.max(
            0,
            toNumber(document.getElementById("delivery-fee").value)
        );

        const orderPayload = {
            number: document.getElementById("order-number").value.trim(),

            date: document.getElementById("order-date").value,

            clientName,

            clientContact: document.getElementById("client-contact")
                .value.trim(),

            deliveryLocation: document.getElementById("delivery-location")
                .value.trim(),

            paymentStatus: document.getElementById("payment-status").value,

            deliveryStatus: document.getElementById("delivery-status").value,

            products,

            subtotal,

            deliveryFee,

            total: subtotal + deliveryFee,

            updatedAt: firebase.firestore.FieldValue.serverTimestamp()
        };

        try {
            if (orderId) {
                await db.collection("orders")
                    .doc(orderId)
                    .update(orderPayload);
            } else {
                orderPayload.createdAt =
                    firebase.firestore.FieldValue.serverTimestamp();

                await db.collection("orders").add(orderPayload);
            }

            alert(t("msgSaved"));

            initForm();

        } catch (error) {
            console.error("Erreur Firestore :", error);

            alert(
                `${t("msgSaveError")}\n${error.message || ""}`
            );
        }
    }
);


// ==========================================================================
// 13. SYNCHRONISATION FIRESTORE EN TEMPS RÉEL
// ==========================================================================

function listenToOrders() {
    if (unsubscribeOrders) {
        unsubscribeOrders();
    }

    unsubscribeOrders = db.collection("orders")
        .onSnapshot(
            snapshot => {
                allOrders = [];

                const clientNames = new Set();

                snapshot.forEach(document => {
                    const data = document.data();

                    const order = {
                        ...data,
                        id: document.id,

                        number: String(data.number ?? ""),
                        date: String(data.date ?? ""),

                        clientName: String(data.clientName ?? ""),
                        clientContact: String(data.clientContact ?? ""),

                        deliveryLocation: String(
                            data.deliveryLocation ?? ""
                        ),

                        paymentStatus: String(
                            data.paymentStatus ?? "unpaid"
                        ),

                        deliveryStatus: String(
                            data.deliveryStatus ?? "new"
                        ),

                        products: normalizeProducts(data.products),

                        subtotal: toNumber(data.subtotal),
                        deliveryFee: toNumber(data.deliveryFee),
                        total: toNumber(data.total)
                    };

                    // Compatibilité avec d'anciennes commandes.
                    if (
                        !Number.isFinite(Number(data.total)) &&
                        order.products.length
                    ) {
                        order.subtotal = order.products.reduce(
                            (sum, product) =>
                                sum + product.price * product.qty,
                            0
                        );

                        order.total =
                            order.subtotal + order.deliveryFee;
                    }

                    allOrders.push(order);

                    if (order.clientName) {
                        clientNames.add(order.clientName);
                    }
                });

                // Tri par date décroissante.
                allOrders.sort((a, b) => {
                    return b.date.localeCompare(a.date);
                });

                updateClientDatalist(clientNames);

                renderTable();

                updateReportsAndPhrases();
            },

            error => {
                console.error("Erreur de lecture Firestore :", error);

                alert(
                    `${t("msgLoadError")}\n${error.message || ""}`
                );
            }
        );
}

listenToOrders();


// ==========================================================================
// 14. AUTOCOMPLÉTION DES CLIENTS
// ==========================================================================

function updateClientDatalist(clientsSet) {
    const datalist = document.getElementById("clients-list");

    if (!datalist) {
        return;
    }

    datalist.innerHTML = "";

    clientsSet.forEach(name => {
        const option = document.createElement("option");

        option.value = name;

        datalist.appendChild(option);
    });
}

document.getElementById("client-name").addEventListener(
    "input",
    event => {
        const inputName = event.target.value.trim().toLowerCase();

        const matchedOrder = allOrders.find(order => {
            return order.clientName.toLowerCase() === inputName;
        });

        if (matchedOrder) {
            document.getElementById("client-contact").value =
                matchedOrder.clientContact;
        }
    }
);


// ==========================================================================
// 15. RECHERCHE ET FILTRES
// ==========================================================================

document.getElementById("filter-search").addEventListener(
    "input",
    renderTable
);

document.getElementById("filter-date").addEventListener(
    "change",
    renderTable
);

document.getElementById("filter-payment").addEventListener(
    "change",
    renderTable
);

document.getElementById("filter-delivery").addEventListener(
    "change",
    renderTable
);

function renderTable() {
    if (!tbody) {
        return;
    }

    const searchInput = document.getElementById("filter-search");
    const dateInput = document.getElementById("filter-date");
    const paymentInput = document.getElementById("filter-payment");
    const deliveryInput = document.getElementById("filter-delivery");

    const search = searchInput.value.trim().toLowerCase();
    const selectedDate = dateInput.value;
    const selectedPayment = paymentInput.value;
    const selectedDelivery = deliveryInput.value;

    tbody.innerHTML = "";

    const filteredOrders = allOrders.filter(order => {
        const searchableText = [
            order.clientName,
            order.clientContact,
            order.number,
            order.deliveryLocation,
            ...order.products.map(product => product.name)
        ].join(" ").toLowerCase();

        const matchesSearch = searchableText.includes(search);

        const matchesDate =
            !selectedDate || order.date === selectedDate;

        const matchesPayment =
            !selectedPayment ||
            order.paymentStatus === selectedPayment;

        const matchesDelivery =
            !selectedDelivery ||
            order.deliveryStatus === selectedDelivery;

        return (
            matchesSearch &&
            matchesDate &&
            matchesPayment &&
            matchesDelivery
        );
    });

    filteredOrders.forEach(order => {
        const tr = document.createElement("tr");

        const productsHTML = order.products.map(product => {
            return `
                <li>
                    ${escapeHTML(product.name)}
                    (x${product.qty})
                </li>
            `;
        }).join("");

        const paymentLabels = {
            unpaid: "optUnpaid",
            partial: "optPartial",
            paid: "optPaid"
        };

        const deliveryLabels = {
            new: "optNew",
            preparing: "optPreparing",
            shipping: "optShipping",
            delivered: "optDelivered"
        };

        const paymentKey =
            paymentLabels[order.paymentStatus] || "optUnpaid";

        const deliveryKey =
            deliveryLabels[order.deliveryStatus] || "optNew";

        const safePaymentStatus =
            /^[a-z]+$/.test(order.paymentStatus)
                ? order.paymentStatus
                : "unpaid";

        const safeDeliveryStatus =
            /^[a-z]+$/.test(order.deliveryStatus)
                ? order.deliveryStatus
                : "new";

        tr.innerHTML = `
            <td>
                <b>${escapeHTML(order.number)}</b>
            </td>

            <td>${escapeHTML(order.date)}</td>

            <td>
                <b>${escapeHTML(order.clientName)}</b>
                <br>
                <small>${escapeHTML(order.clientContact)}</small>
            </td>

            <td>${escapeHTML(order.deliveryLocation)}</td>

            <td>
                <ul class="table-products-list">
                    ${productsHTML}
                </ul>
            </td>

            <td>${formatMoney(order.subtotal)} $</td>

            <td>${formatMoney(order.deliveryFee)} $</td>

            <td>
                <mark>
                    <b>${formatMoney(order.total)} $</b>
                </mark>
            </td>

            <td>
                <span class="badge badge-${safePaymentStatus}">
                    ${escapeHTML(t(paymentKey))}
                </span>
            </td>

            <td>
                <span class="badge badge-${safeDeliveryStatus}">
                    ${escapeHTML(t(deliveryKey))}
                </span>
            </td>

            <td>
                <div class="action-btn-group">
                    <button
                        type="button"
                        class="btn btn-secondary btn-sm"
                        data-action="edit"
                        data-id="${escapeHTML(order.id)}"
                        aria-label="Edit order"
                    >
                        ✎
                    </button>

                    <button
                        type="button"
                        class="btn btn-danger btn-sm"
                        data-action="delete"
                        data-id="${escapeHTML(order.id)}"
                        aria-label="Delete order"
                    >
                        ✕
                    </button>
                </div>
            </td>
        `;

        tbody.appendChild(tr);
    });
}


// ==========================================================================
// 16. ACTIONS DU TABLEAU
// ==========================================================================

tbody.addEventListener("click", event => {
    const button = event.target.closest("button[data-action]");

    if (!button) {
        return;
    }

    const id = button.dataset.id;

    if (button.dataset.action === "edit") {
        window.editOrder(id);
    }

    if (button.dataset.action === "delete") {
        window.deleteOrder(id);
    }
});


// ==========================================================================
// 17. MODIFICATION D'UNE COMMANDE
// ==========================================================================

window.editOrder = function(id) {
    const order = allOrders.find(item => item.id === id);

    if (!order) {
        return;
    }

    document.getElementById("order-id").value = order.id;

    document.getElementById("order-number").value = order.number;

    document.getElementById("order-date").value = order.date;

    document.getElementById("client-name").value = order.clientName;

    document.getElementById("client-contact").value =
        order.clientContact;

    document.getElementById("delivery-location").value =
        order.deliveryLocation;

    document.getElementById("payment-status").value =
        order.paymentStatus;

    document.getElementById("delivery-status").value =
        order.deliveryStatus;

    document.getElementById("delivery-fee").value =
        order.deliveryFee;

    productsContainer.innerHTML = "";

    if (order.products.length > 0) {
        order.products.forEach(product => {
            createProductRow(product);
        });
    } else {
        createProductRow();
    }

    calculateOrderTotals();

    document.getElementById("btn-cancel-edit")
        .classList.remove("hidden");

    const title = document.getElementById("form-title");

    title.setAttribute("data-i18n", "editOrderTitle");
    title.textContent = t("editOrderTitle");

    window.scrollTo({
        top: 0,
        behavior: "smooth"
    });
};

document.getElementById("btn-cancel-edit").addEventListener(
    "click",
    initForm
);


// ==========================================================================
// 18. SUPPRESSION D'UNE COMMANDE
// ==========================================================================

window.deleteOrder = async function(id) {
    const confirmed = confirm(t("msgConfirmDelete"));

    if (!confirmed) {
        return;
    }

    try {
        await db.collection("orders").doc(id).delete();

        alert(t("msgDeleted"));

    } catch (error) {
        console.error("Erreur de suppression :", error);

        alert(
            `${t("msgDeleteError")}\n${error.message || ""}`
        );
    }
};


// ==========================================================================
// 19. TABLEAU DE BORD ET RAPPORT JOURNALIER
// ==========================================================================

function updateReportsAndPhrases() {
    const totalOrders = allOrders.length;

    const totalRevenue = allOrders.reduce(
        (sum, order) => sum + order.total,
        0
    );

    const pendingDeliveries = allOrders.filter(order => {
        return order.deliveryStatus !== "delivered";
    }).length;

    document.getElementById("stat-total-orders").textContent =
        totalOrders;

    document.getElementById("stat-total-revenue").textContent =
        `${formatMoney(totalRevenue)} $`;

    document.getElementById("stat-pending-delivery").textContent =
        pendingDeliveries;

    const today = getLocalDate();

    const todayOrders = allOrders.filter(order => {
        return order.date === today;
    });

    const reportElement = document.getElementById(
        "daily-report-phrase"
    );

    if (todayOrders.length === 0) {
        reportElement.textContent = t("noDataToday");

        return;
    }

    const todayRevenue = todayOrders.reduce(
        (sum, order) => sum + order.total,
        0
    );

    const todayPending = todayOrders.filter(order => {
        return order.deliveryStatus !== "delivered";
    }).length;

    const report = t("phraseReport")
        .replace("{count}", todayOrders.length)
        .replace("{total}", formatMoney(todayRevenue))
        .replace("{pending}", todayPending);

    reportElement.textContent = report;
}


// ==========================================================================
// 20. EXPORTATION EXCEL AVEC SHEETJS
// ==========================================================================

document.getElementById("btn-export-excel").addEventListener(
    "click",
    () => {
        if (allOrders.length === 0) {
            alert(t("msgNoOrders"));
            return;
        }

        if (typeof XLSX === "undefined") {
            alert(t("msgExportError"));
            return;
        }

        try {
            const rows = allOrders.map(order => ({
                "Numéro Commande": order.number,

                "Date": order.date,

                "Nom Client": order.clientName,

                "Contact": order.clientContact,

                "Lieu Livraison": order.deliveryLocation,

                "Produits": order.products.map(product => {
                    return `${product.name} (x${product.qty})`;
                }).join(", "),

                "Sous-Total ($)": order.subtotal,

                "Frais Livraison ($)": order.deliveryFee,

                "Total Général ($)": order.total,

                "Statut Paiement": order.paymentStatus.toUpperCase(),

                "Statut Logistique": order.deliveryStatus.toUpperCase()
            }));

            const worksheet = XLSX.utils.json_to_sheet(rows);

            // Ajuster automatiquement la largeur des colonnes.
            worksheet["!cols"] = Object.keys(rows[0]).map(key => {
                const longestValue = Math.max(
                    key.length,
                    ...rows.map(row => String(row[key] ?? "").length)
                );

                return {
                    wch: Math.min(longestValue + 2, 45)
                };
            });

            const workbook = XLSX.utils.book_new();

            XLSX.utils.book_append_sheet(
                workbook,
                worksheet,
                "Registre Commandes"
            );

            XLSX.writeFile(
                workbook,
                `Registre_Commandes_${getLocalDate()}.xlsx`
            );

        } catch (error) {
            console.error("Erreur d'exportation :", error);

            alert(t("msgExportError"));
        }
    }
);


// ==========================================================================
// 21. INITIALISATION DE L'APPLICATION
// ==========================================================================

initForm();
