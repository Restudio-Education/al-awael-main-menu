/* ======================================================
   📬 Microsoft Graph API - الإشعارات الكاملة
   Al-Awael Platform - Restudio Education
   ====================================================== */

const GRAPH_CONFIG = {
    clientId: 'f43f61b1-031b-4b40-9305-70e11d57459e',
    tenantId: '260eb6b4-9eda-4872-b458-fc4bdde5ddd4',
    redirectUri: 'https://restudio-education.github.io/al-awael-main-menu/',
    scopes: ['User.Read', 'Mail.Read', 'Calendars.Read'],
    graphEndpoint: 'https://graph.microsoft.com/v1.0'
};

let msalInstance = null;
let msalInitialized = false;

/* ===== تهيئة MSAL v3 ===== */
async function initMSAL() {
    if (typeof msal === 'undefined') {
        console.error('❌ MSAL.js not loaded');
        return null;
    }
    if (msalInstance && msalInitialized) return msalInstance;

    msalInstance = new msal.PublicClientApplication({
        auth: {
            clientId: GRAPH_CONFIG.clientId,
            authority: `https://login.microsoftonline.com/${GRAPH_CONFIG.tenantId}`,
            redirectUri: GRAPH_CONFIG.redirectUri,
            navigateToLoginRequestUrl: false
        },
        cache: {
            cacheLocation: 'localStorage',
            storeAuthStateInCookie: false
        }
    });

    await msalInstance.initialize();
    msalInitialized = true;
    console.log('✅ MSAL initialized');
    return msalInstance;
}

/* ===== تسجيل الدخول ===== */
async function graphSignIn() {
    if (!msalInitialized) await initMSAL();
    const loginRequest = { scopes: GRAPH_CONFIG.scopes, prompt: 'select_account' };
    const response = await msalInstance.loginPopup(loginRequest);
    console.log('✅ Login success:', response.account.username);
    return response;
}

/* ===== تسجيل الخروج ===== */
async function graphSignOut() {
    if (!msalInitialized) await initMSAL();
    const accounts = msalInstance.getAllAccounts();
    if (accounts.length > 0) {
        await msalInstance.logoutPopup({
            account: accounts[0],
            postLogoutRedirectUri: GRAPH_CONFIG.redirectUri
        });
    }
    console.log('✅ Logged out');
}

/* ===== الحصول على Token ===== */
async function getGraphToken() {
    if (!msalInitialized) await initMSAL();
    const accounts = msalInstance.getAllAccounts();
    if (accounts.length === 0) return null;

    const silentRequest = { scopes: GRAPH_CONFIG.scopes, account: accounts[0] };

    try {
        const response = await msalInstance.acquireTokenSilent(silentRequest);
        return response.accessToken;
    } catch (error) {
        if (error instanceof msal.InteractionRequiredAuthError) {
            try {
                const response = await msalInstance.acquireTokenPopup(silentRequest);
                return response.accessToken;
            } catch (e) { return null; }
        }
        return null;
    }
}

/* ===== معلومات المستخدم ===== */
async function getGraphUserInfo(token) {
    try {
        const res = await fetch(`${GRAPH_CONFIG.graphEndpoint}/me`, {
            headers: { 'Authorization': `Bearer ${token}` }
        });
        if (!res.ok) throw new Error(`HTTP ${res.status}`);
        return await res.json();
    } catch (err) {
        console.error('❌ Get user info failed:', err);
        return null;
    }
}

/* ===== جلب الرسائل ===== */
async function fetchGraphMessages(token, top = 15) {
    try {
        const url = `${GRAPH_CONFIG.graphEndpoint}/me/messages?$top=${top}&$select=id,subject,from,receivedDateTime,isRead,bodyPreview,webLink&$orderby=receivedDateTime DESC`;
        const res = await fetch(url, {
            headers: { 'Authorization': `Bearer ${token}` }
        });
        if (!res.ok) throw new Error(`HTTP ${res.status}`);
        const data = await res.json();
        console.log(`📧 Messages fetched: ${data.value?.length || 0}`);
        return data.value || [];
    } catch (err) {
        console.error('❌ Fetch messages failed:', err);
        return [];
    }
}

/* ===== جلب المواعيد ===== */
async function fetchGraphEvents(token, top = 5) {
    try {
        const now = new Date().toISOString();
        const url = `${GRAPH_CONFIG.graphEndpoint}/me/events?$top=${top}&$filter=start/dateTime ge '${now}'&$orderby=start/dateTime&$select=id,subject,start,end,location,webLink`;
        const res = await fetch(url, {
            headers: { 'Authorization': `Bearer ${token}` }
        });
        if (!res.ok) return [];
        const data = await res.json();
        return data.value || [];
    } catch (err) {
        console.error('❌ Fetch events failed:', err);
        return [];
    }
}

/* ===== تصدير API ===== */
window.GraphAPI = {
    init: initMSAL,
    signIn: graphSignIn,
    signOut: graphSignOut,
    getToken: getGraphToken,
    getUserInfo: getGraphUserInfo,
    fetchMessages: fetchGraphMessages,
    fetchEvents: fetchGraphEvents,
    isLoggedIn: () => {
        if (!msalInstance) return false;
        return msalInstance.getAllAccounts().length > 0;
    },
    getAccount: () => {
        if (!msalInstance) return null;
        const accounts = msalInstance.getAllAccounts();
        return accounts.length > 0 ? accounts[0] : null;
    }
};

/* ======================================================
   🎨 عرض الإشعارات
   ====================================================== */
async function loadRealNotifications() {
    console.log('📬 loadRealNotifications started');
    const modalContent = document.querySelector('#notificationModal .modal-content');
    if (!modalContent) {
        console.error('❌ modal-content not found');
        return;
    }

    modalContent.innerHTML = `
        <div class="modal-logo">
            <img src="https://i.postimg.cc/bJ4RsLjj/email-(2).png" alt="إشعارات">
        </div>
        <div class="notif-loading">جاري تحميل الإشعارات من Microsoft...</div>
    `;

    try {
        await initMSAL();
    } catch (e) {
        console.error('❌ MSAL init failed:', e);
        modalContent.innerHTML = `
            <div class="modal-logo"><img src="https://i.postimg.cc/bJ4RsLjj/email-(2).png" alt="إشعارات"></div>
            <div class="notif-error">تعذّر تهيئة نظام الإشعارات.</div>
            <button class="close-modal" id="closeNotificationModalBtn">إغلاق</button>
        `;
        document.getElementById('closeNotificationModalBtn')
            ?.addEventListener('click', () => {
                document.getElementById('notificationModal').style.display = 'none';
            });
        return;
    }

    if (!window.GraphAPI.isLoggedIn()) {
        console.log('⚠️ Not logged in');
        modalContent.innerHTML = `
            <div class="modal-logo">
                <img src="https://i.postimg.cc/bJ4RsLjj/email-(2).png" alt="إشعارات">
            </div>
            <div class="support-text">
                لعرض إشعارات بريدك، سجّل الدخول بحساب Microsoft الخاص بمدرستك.
            </div>
            <button id="graphSignInBtn" class="notif-signin-btn">
                <i class="fas fa-windows"></i> تسجيل الدخول بـ Microsoft
            </button>
            <div class="notification-note">
                <i class="fas fa-shield-alt"></i>
                بياناتك آمنة ومحمية — لا نحتفظ بها
            </div>
            <button class="close-modal" id="closeNotificationModalBtn">إغلاق</button>
        `;
        document.getElementById('graphSignInBtn')?.addEventListener('click', async (ev) => {
            const btn = ev.currentTarget;
            btn.disabled = true;
            btn.textContent = 'جاري التهيئة...';
            try {
                if (!msalInitialized) await initMSAL();
                await window.GraphAPI.signIn();
                loadRealNotifications();
            } catch (e) {
                console.error('Sign in failed:', e);
                alert('تعذر تسجيل الدخول: ' + (e.message || ''));
                btn.disabled = false;
                btn.innerHTML = '<i class="fas fa-windows"></i> تسجيل الدخول بـ Microsoft';
            }
        });
        document.getElementById('closeNotificationModalBtn')
            ?.addEventListener('click', () => {
                document.getElementById('notificationModal').style.display = 'none';
            });
        return;
    }

    try {
        console.log('🔑 Getting token...');
        const token = await window.GraphAPI.getToken();
        if (!token) {
            modalContent.innerHTML = `
                <div class="modal-logo"><img src="https://i.postimg.cc/bJ4RsLjj/email-(2).png" alt="إشعارات"></div>
                <div class="support-text">انتهت الجلسة. سجّل الدخول من جديد.</div>
                <button class="close-modal" id="closeNotificationModalBtn">إغلاق</button>
            `;
            document.getElementById('closeNotificationModalBtn')
                ?.addEventListener('click', () => {
                    document.getElementById('notificationModal').style.display = 'none';
                });
            return;
        }

        console.log('📨 Fetching messages & events...');
        const [user, messages, events] = await Promise.all([
            window.GraphAPI.getUserInfo(token),
            window.GraphAPI.fetchMessages(token, 15),
            window.GraphAPI.fetchEvents(token, 5)
        ]);

        const unread = messages.filter(m => !m.isRead).length;
        console.log(`✅ Got ${messages.length} messages, ${unread} unread, ${events.length} events`);

        // ✅ تحديث شارة الجرس بالعدد الحقيقي
        const badge = document.getElementById('notificationCount');
        if (badge) {
            if (unread > 0) {
                badge.textContent = unread > 99 ? '99+' : String(unread);
                badge.style.display = 'flex';
            } else {
                badge.style.display = 'none';
            }
        }

        let html = `
            <div class="modal-logo">
                <img src="https://i.postimg.cc/bJ4RsLjj/email-(2).png" alt="إشعارات">
            </div>
            <div class="notif-header">
                <div class="notif-greeting">مرحباً ${user?.displayName || 'عزيزي'} 👋</div>
                <div class="notif-stats">
                    <strong>${unread}</strong> غير مقروءة • 
                    <strong>${messages.length}</strong> إجمالي • 
                    <strong>${events.length}</strong> موعد
                </div>
            </div>
        `;

        if (messages.length === 0) {
            html += `<div class="notif-empty">لا توجد رسائل حديثة.</div>`;
        } else {
            html += `<div class="notif-list">`;
            messages.forEach(m => {
                const from = m.from?.emailAddress?.name || m.from?.emailAddress?.address || 'غير معروف';
                const date = new Date(m.receivedDateTime).toLocaleString('ar-SA', {
                    day: 'numeric', month: 'short', hour: '2-digit', minute: '2-digit'
                });
                const cls = m.isRead ? 'read' : 'unread';
                html += `
                    <a href="${m.webLink}" target="_blank" rel="noopener noreferrer" class="notif-item ${cls}">
                        <div class="notif-subject">${m.subject || '(بدون عنوان)'}</div>
                        <div class="notif-meta">من: ${from} • ${date}</div>
                    </a>
                `;
            });
            html += `</div>`;
        }

        if (events.length > 0) {
            html += `
                <div class="notif-events-block">
                    <div class="notif-events-title">المواعيد القادمة</div>
            `;
            events.slice(0, 5).forEach(ev => {
                const start = new Date(ev.start.dateTime).toLocaleString('ar-SA', {
                    day: 'numeric', month: 'short', hour: '2-digit', minute: '2-digit'
                });
                html += `
                    <a href="${ev.webLink}" target="_blank" rel="noopener noreferrer" class="notif-event-item">
                        <strong>${ev.subject || '(بدون عنوان)'}</strong>
                        ${start} ${ev.location?.displayName ? '• ' + ev.location.displayName : ''}
                    </a>
                `;
            });
            html += `</div>`;
        }

        html += `
            <button class="close-modal" id="closeNotificationModalBtn">إغلاق</button>
            <button class="notif-logout-btn" id="graphLogoutBtn">
                <i class="fas fa-sign-out-alt"></i> تسجيل الخروج
            </button>
        `;

        modalContent.innerHTML = html;

        document.getElementById('closeNotificationModalBtn')
            ?.addEventListener('click', () => {
                document.getElementById('notificationModal').style.display = 'none';
            });

        document.getElementById('graphLogoutBtn')
            ?.addEventListener('click', async () => {
                await window.GraphAPI.signOut();
                document.getElementById('notificationModal').style.display = 'none';
                const badge = document.getElementById('notificationCount');
                if (badge) badge.style.display = 'none';
            });

    } catch (err) {
        console.error('❌ Graph notifications error:', err);
        modalContent.innerHTML = `
            <div class="modal-logo"><img src="https://i.postimg.cc/bJ4RsLjj/email-(2).png" alt="إشعارات"></div>
            <div class="notif-error">تعذّر تحميل الإشعارات.<br><small>${err.message || ''}</small></div>
            <button class="close-modal" id="closeNotificationModalBtn">إغلاق</button>
        `;
        document.getElementById('closeNotificationModalBtn')
            ?.addEventListener('click', () => {
                document.getElementById('notificationModal').style.display = 'none';
            });
    }
}

window.loadRealNotifications = loadRealNotifications;

/* ===== ربط زر الإشعارات ===== */
function bindNotificationButton() {
    const btn = document.getElementById('notificationBtn');
    if (!btn) {
        console.warn('⚠️ notificationBtn not found');
        return;
    }
    if (btn._graphBound) return;
    btn._graphBound = true;
    btn.addEventListener('click', function() {
        console.log('🔔 Bell clicked!');
        const modal = document.getElementById('notificationModal');
        if (modal) modal.style.display = 'flex';
        window.loadRealNotifications();
    });
    console.log('✅ Notification button bound');
}

/* ===== التهيئة عند تحميل الصفحة ===== */
async function bootstrapMSAL() {
    try {
        await initMSAL();
        console.log('✅ MSAL bootstrap complete');
    } catch (e) {
        console.error('❌ MSAL bootstrap failed:', e);
    }
    bindNotificationButton();
}

if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', bootstrapMSAL);
} else {
    bootstrapMSAL();
}

console.log('✅ Graph API module loaded');
