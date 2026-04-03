const functions = require("firebase-functions");
const admin = require("firebase-admin");
const bcrypt = require("bcryptjs");
const cors = require("cors")({ origin: true });

admin.initializeApp();
const db = admin.firestore();

// ---------- helpers ----------
function nowTs() {
    return admin.firestore.Timestamp.now();
}

async function nextCounter(name) {
    const ref = db.collection("counters").doc(name);
    return await db.runTransaction(async (tx) => {
        const snap = await tx.get(ref);
        const cur = snap.exists ? (snap.data().next || 1) : 1;
        tx.set(ref, { next: cur + 1 }, { merge: true });
        return cur;
    });
}

function requireAuth(req) {
    const user = req.user; // set by verifyIdToken middleware (below)
    if (!user) throw new Error("UNAUTHENTICATED");
    return user;
}

function requireApp(user, appId) {
    const allowed = user.allowedApps || [];
    const ok = allowed.includes("*") || allowed.includes(appId);
    if (!ok) throw new Error("FORBIDDEN_APP");
}

function requireMinAccess(user, minLevel) {
    const lvl = Number(user.accessLevel || 0);
    if (lvl < Number(minLevel || 0)) throw new Error("FORBIDDEN_LEVEL");
}

// ---------- middleware: verify Firebase ID token ----------
async function verifyIdToken(req) {
    const auth = req.headers.authorization || "";
    const match = auth.match(/^Bearer (.+)$/);
    if (!match) return null;
    const idToken = match[1];
    try {
        const decoded = await admin.auth().verifyIdToken(idToken);
        return decoded; // contains custom claims
    } catch (e) {
        return null;
    }
}

// ---------- HTTPS API ----------
exports.api = functions.https.onRequest((req, res) => {
    cors(req, res, async () => {
        try {
            // attach user if token exists
            req.user = await verifyIdToken(req);

            // ROUTING
            const path = (req.path || "/").replace(/\/+$/, "");
            const method = req.method.toUpperCase();

            // --- AUTH: username + PIN -> custom token ---
            if (method === "POST" && path === "/auth/login") {
                const { username, pin } = req.body || {};
                const u = String(username || "").trim().toLowerCase();
                const p = String(pin || "").trim();

                if (!u || !p) return res.status(400).json({ ok: false, error: "VALIDATION_ERROR" });

                const snap = await db.collection("employees")
                    .where("username", "==", u)
                    .limit(1)
                    .get();

                if (snap.empty) return res.status(401).json({ ok: false, error: "INVALID_LOGIN" });

                const emp = snap.docs[0].data();

                // status + lock checks
                if (String(emp.status || "").toLowerCase() !== "active") {
                    return res.status(403).json({ ok: false, error: "DISABLED" });
                }
                if (emp.lockedUntil && emp.lockedUntil.toMillis && emp.lockedUntil.toMillis() > Date.now()) {
                    return res.status(403).json({ ok: false, error: "LOCKED", lockedUntil: emp.lockedUntil });
                }

                const ok = await bcrypt.compare(p, String(emp.pinHash || ""));
                const empRef = db.collection("employees").doc(emp.employeeId);

                if (!ok) {
                    const failed = Number(emp.failedAttempts || 0) + 1;
                    const MAX = 5;
                    let lockedUntil = null;
                    if (failed >= MAX) {
                        lockedUntil = admin.firestore.Timestamp.fromMillis(Date.now() + 15 * 60 * 1000);
                    }
                    await empRef.set({ failedAttempts: failed, lockedUntil, updatedAt: nowTs() }, { merge: true });
                    return res.status(401).json({ ok: false, error: "INVALID_LOGIN" });
                }

                // reset attempts
                await empRef.set({ failedAttempts: 0, lockedUntil: null, updatedAt: nowTs() }, { merge: true });

                // custom claims (keep small)
                const claims = {
                    employeeId: emp.employeeId,
                    username: emp.username,
                    name: emp.name || "",
                    role: emp.role || "",
                    accessLevel: Number(emp.accessLevel || 0),
                    allowedApps: emp.allowedApps || []
                };

                // We need a Firebase Auth user UID to mint token. We'll use employeeId as UID.
                // Create user if not exists.
                try {
                    await admin.auth().getUser(emp.employeeId);
                } catch (e) {
                    await admin.auth().createUser({
                        uid: emp.employeeId,
                        displayName: emp.name || emp.username,
                        disabled: false
                    });
                }
                await admin.auth().setCustomUserClaims(emp.employeeId, claims);
                const token = await admin.auth().createCustomToken(emp.employeeId);

                return res.json({ ok: true, token, employee: claims });
            }

            // from here on: require Firebase ID token
            const user = requireAuth(req);

            // --- APPS: list allowed apps ---
            if (method === "GET" && path === "/apps") {
                // If allowedApps=["*"], return all visible active apps
                const allowed = user.allowedApps || [];
                let q = db.collection("apps").where("active", "==", true).where("visible", "==", true);
                const appsSnap = await q.get();
                let apps = appsSnap.docs.map(d => d.data());

                if (!(allowed.includes("*"))) {
                    apps = apps.filter(a => allowed.includes(a.appId));
                }

                // also enforce minAccessLevel
                apps = apps.filter(a => Number(user.accessLevel || 0) >= Number(a.minAccessLevel || 0));
                return res.json({ ok: true, apps });
            }

            // --- EMPLOYEES (admin app required) ---
            if (path.startsWith("/employees")) {
                requireApp(user, "EMP_ADMIN");
                requireMinAccess(user, 9);

                if (method === "GET" && path === "/employees") {
                    const snap = await db.collection("employees").orderBy("employeeId").get();
                    const employees = snap.docs.map(d => {
                        const e = d.data();
                        delete e.pinHash;
                        return e;
                    });
                    return res.json({ ok: true, employees });
                }

                if (method === "POST" && path === "/employees") {
                    const body = req.body || {};
                    const employeeId = String(body.employeeId || "").trim();
                    const username = String(body.username || "").trim().toLowerCase();
                    if (!employeeId || !username) return res.status(400).json({ ok: false, error: "VALIDATION_ERROR" });

                    // if pin provided, hash it
                    let patch = {
                        employeeId,
                        name: body.name || "",
                        username,
                        email: body.email || "",
                        phone: body.phone || "",
                        role: body.role || "",
                        accessLevel: Number(body.accessLevel || 0),
                        allowedApps: Array.isArray(body.allowedApps) ? body.allowedApps : [],
                        status: String(body.status || "active").toLowerCase(),
                        updatedAt: nowTs(),
                    };

                    if (body.pin) {
                        patch.pinHash = await bcrypt.hash(String(body.pin), 10);
                    }

                    // ensure createdAt
                    const ref = db.collection("employees").doc(employeeId);
                    const existing = await ref.get();
                    if (!existing.exists) patch.createdAt = nowTs();

                    await ref.set(patch, { merge: true });
                    return res.json({ ok: true });
                }

                if (method === "POST" && path.match(/^\/employees\/[^/]+\/reset-pin$/)) {
                    const employeeId = path.split("/")[2];
                    const { pin } = req.body || {};
                    if (!pin) return res.status(400).json({ ok: false, error: "VALIDATION_ERROR" });
                    const ref = db.collection("employees").doc(employeeId);
                    await ref.set({
                        pinHash: await bcrypt.hash(String(pin), 10),
                        failedAttempts: 0,
                        lockedUntil: null,
                        updatedAt: nowTs()
                    }, { merge: true });
                    return res.json({ ok: true });
                }
            }

            // --- INCOME / EXPENSE ---
            if (path.startsWith("/income-expense")) {
                requireApp(user, "INCOME_EXPENSE");

                if (method === "GET" && path === "/income-expense") {
                    const from = req.query.from ? new Date(req.query.from) : null;
                    const to = req.query.to ? new Date(req.query.to) : null;
                    const status = (req.query.status || "ACTIVE").toUpperCase();

                    let q = db.collection("incomeExpense").where("status", "==", status).orderBy("ts", "desc");
                    if (from) q = q.where("ts", ">=", admin.firestore.Timestamp.fromDate(from));
                    if (to) q = q.where("ts", "<=", admin.firestore.Timestamp.fromDate(to));

                    const snap = await q.limit(500).get();
                    return res.json({ ok: true, rows: snap.docs.map(d => d.data()) });
                }

                if (method === "POST" && path === "/income-expense") {
                    const b = req.body || {};
                    const entryId = await nextCounter("incomeExpense");
                    const docId = String(entryId).padStart(6, "0");

                    const data = {
                        entryId,
                        ts: nowTs(),
                        employeeId: user.employeeId,
                        employeeName: user.name || "",
                        username: user.username || "",
                        nature: b.nature || "",
                        division: b.division || "",
                        cashTransfer: b.cashTransfer || "",
                        invoice: b.invoice || "",
                        type: b.type || "",
                        description: b.description || "",
                        amount: Number(b.amount || 0),
                        remark: b.remark || "",
                        status: "ACTIVE"
                    };

                    await db.collection("incomeExpense").doc(docId).set(data);
                    return res.json({ ok: true, entryId });
                }

                if (method === "POST" && path.match(/^\/income-expense\/\d+\/void$/)) {
                    const entryId = Number(path.split("/")[2]);
                    const docId = String(entryId).padStart(6, "0");
                    const reason = (req.body && req.body.reason) || "";

                    await db.collection("incomeExpense").doc(docId).set({
                        status: "VOID",
                        voidInfo: { voidBy: user.username, voidAt: nowTs(), reason }
                    }, { merge: true });

                    return res.json({ ok: true });
                }
            }

            // --- BANK MASTERS + ACCOUNTS ---
            if (path.startsWith("/bank")) {
                requireApp(user, "BANK");

                if (method === "GET" && path === "/bank/masters") {
                    const snap = await db.collection("bankMasters").where("active", "==", true).orderBy("bankName").get();
                    return res.json({ ok: true, rows: snap.docs.map(d => ({ id: d.id, ...d.data() })) });
                }

                if (method === "POST" && path === "/bank/masters") {
                    requireMinAccess(user, 5);
                    const { bankName, type } = req.body || {};
                    if (!bankName || !type) return res.status(400).json({ ok: false, error: "VALIDATION_ERROR" });
                    const id = String(bankName).trim().toLowerCase().replace(/\s+/g, "_");
                    await db.collection("bankMasters").doc(id).set({ bankName, type, active: true }, { merge: true });
                    return res.json({ ok: true });
                }

                if (method === "GET" && path === "/bank/accounts") {
                    const snap = await db.collection("bankAccounts").orderBy("no", "desc").limit(500).get();
                    return res.json({ ok: true, rows: snap.docs.map(d => d.data()) });
                }

                if (method === "POST" && path === "/bank/accounts") {
                    requireMinAccess(user, 5);
                    const b = req.body || {};
                    const no = await nextCounter("bankAccounts");
                    const docId = String(no).padStart(6, "0");
                    const data = {
                        no,
                        accountType: b.accountType || "",
                        accountUserName: b.accountUserName || "",
                        status: b.status || "Active",
                        accountNumber: b.accountNumber || "",
                        nrc: b.nrc || "",
                        registerPhoneNumber: b.registerPhoneNumber || "",
                        openingDate: b.openingDate ? admin.firestore.Timestamp.fromDate(new Date(b.openingDate)) : null,
                        closedDate: null,
                        createdAt: nowTs(),
                        updatedAt: nowTs()
                    };
                    await db.collection("bankAccounts").doc(docId).set(data);
                    return res.json({ ok: true, no });
                }

                if (method === "POST" && path.match(/^\/bank\/accounts\/\d+\/close$/)) {
                    requireMinAccess(user, 5);
                    const no = Number(path.split("/")[3]);
                    const docId = String(no).padStart(6, "0");
                    await db.collection("bankAccounts").doc(docId).set({
                        status: "Closed",
                        closedDate: nowTs(),
                        updatedAt: nowTs()
                    }, { merge: true });
                    return res.json({ ok: true });
                }
            }

            // --- DEPOSIT / WITHDRAW ---
            if (path.startsWith("/dw")) {
                requireApp(user, "DW");

                if (method === "GET" && path === "/dw") {
                    const from = req.query.from ? new Date(req.query.from) : null;
                    const to = req.query.to ? new Date(req.query.to) : null;
                    const status = (req.query.status || "ACTIVE").toUpperCase();

                    let q = db.collection("depositWithdraw").where("status", "==", status).orderBy("ts", "desc");
                    if (from) q = q.where("ts", ">=", admin.firestore.Timestamp.fromDate(from));
                    if (to) q = q.where("ts", "<=", admin.firestore.Timestamp.fromDate(to));
                    const snap = await q.limit(500).get();
                    return res.json({ ok: true, rows: snap.docs.map(d => d.data()) });
                }

                if (method === "POST" && path === "/dw") {
                    const b = req.body || {};
                    const entryId = await nextCounter("depositWithdraw");
                    const docId = String(entryId).padStart(6, "0");

                    const data = {
                        entryId,
                        ts: nowTs(),
                        nature: b.nature || "",
                        bankAccount: b.bankAccount || "",
                        selectBankAccount: b.selectBankAccount || "",
                        userAccountName: b.userAccountName || "",
                        amount: Number(b.amount || 0),
                        processedBy: user.username || "",
                        promotion: !!b.promotion,
                        commentRemark: b.commentRemark || "",
                        status: "ACTIVE",
                        sourceEntryId: b.sourceEntryId ? Number(b.sourceEntryId) : null
                    };

                    await db.collection("depositWithdraw").doc(docId).set(data);
                    return res.json({ ok: true, entryId });
                }

                if (method === "POST" && path.match(/^\/dw\/\d+\/void$/)) {
                    const entryId = Number(path.split("/")[2]);
                    const docId = String(entryId).padStart(6, "0");
                    const reason = (req.body && req.body.reason) || "";

                    await db.collection("depositWithdraw").doc(docId).set({
                        status: "VOID",
                        voidInfo: { voidBy: user.username, voidAt: nowTs(), reason }
                    }, { merge: true });

                    return res.json({ ok: true });
                }
            }

            return res.status(404).json({ ok: false, error: "NOT_FOUND" });
        } catch (err) {
            const msg = String(err && err.message ? err.message : err);
            const code =
                msg.includes("UNAUTHENTICATED") ? 401 :
                    msg.includes("FORBIDDEN") ? 403 :
                        500;
            return res.status(code).json({ ok: false, error: msg });
        }
    });
});
