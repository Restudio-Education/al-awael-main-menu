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

/* ===== تهيئة MSAL (v3 - async) ===== */
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

    // ✅ MSAL v3: يجب استدعاء initialize() أولاً
    await msalInstance.initialize();
    msalInitialized = true;

    console.log('✅ MSAL initialized');
    return msalInstance;
}

/* ===== تسجيل الدخول ===== */
async function graphSignIn() {
    if (!msalInitialized) await initMSAL();
    const loginRequest = {
        scopes: GRAPH_CONFIG.scopes,
        prompt: 'select_account'
    };
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

    const silentRequest = {
        scopes: GRAPH_CONFIG.scopes,
        account: accounts[0]
    };

    try {
        const response = await msalInstance.acquireTokenSilent(silentRequest);
        return response.accessToken;
    } catch (error) {
        if (error instanceof msal.InteractionRequiredAuthError) {
            try {
                const response = await msalInstance.acquireTokenPopup(silentRequest);
                return response.accessToken;
            } catch (e) {
                return null;
            }
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
        if (!res.ok) throw new Error(`HTTP ${res.status}`);
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
        <div class="support-text">جاري تحميل الإشعارات من Microsoft...</div>
        <div class="skeleton-text" style="height:60px;margin:12px 0;"></div>
        <div class="skeleton-text" style="height:60px;margin:12px 0;"></div>
    `;

    try {
        // ✅ تهيئة MSAL أولاً
        await initMSAL();
    } catch (e) {
        console.error('❌ MSAL init failed:', e);
        modalContent.innerHTML = `
            <div class="modal-logo"><img src="https://i.postimg.cc/bJ4RsLjj/email-(2).png" alt="إشعارات"></div>
            <div class="support-text">تعذّر تهيئة نظام الإشعارات.<br><small style="color:#64748b;font-size:0.75rem;">${e.message || ''}</small></div>
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
            <button id="graphSignInBtn" class="outlook-link-btn" style="cursor:pointer;border:none;font-family:inherit;">
                <i class="ti ti-brand-windows"></i> تسجيل الدخول بـ Microsoft
            </button>
            <div class="notification-note">
                <i class="ti ti-shield-lock"></i>
                بياناتك آمنة ومحمية — لا نحتفظ بها
            </div>
            <button class="close-modal" id="closeNotificationModalBtn">إغلاق</button>
        `;
        document.getElementById('graphSignInBtn')?.addEventListener('click', async () => {
            try {
                await window.GraphAPI.signIn();
                loadRealNotifications();
            } catch (e) {
                console.error('Sign in failed:', e);
                alert('تعذر تسجيل الدخول: ' + (e.message || ''));
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

        const badge = document.querySelector('.notification-count');
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
            <div class="support-text" style="margin-bottom:14px;">
                مرحباً <strong>${user?.displayName || 'عزيزي'}</strong> 👋<br>
                <span style="font-size:0.85rem;color:#64748b;">
                    <strong>${unread}</strong> رسالة غير مقروءة • <strong>${events.length}</strong> موعد قادم
                </span>
            </div>
        `;

        if (messages.length === 0) {
            html += `<p class="notification-note">لا توجد رسائل حديثة.</p>`;
        } else {
            html += `<div style="max-height:320px;overflow-y:auto;text-align:right;margin:10px 0;">`;
            messages.forEach(m => {
                const from = m.from?.emailAddress?.name || m.from?.emailAddress?.address || 'غير معروف';
                const date = new Date(m.receivedDateTime).toLocaleString('ar-SA', {
                    day: 'numeric', month: 'short', hour: '2-digit', minute: '2-digit'
                });
                html += `
                    <a href="${m.webLink}" target="_blank" rel="noopener noreferrer"
                       style="display:block;padding:10px 12px;margin-bottom:6px;border-radius:10px;
                              background:${m.isRead ? '#f8fafc' : '#eff6ff'};
                              border-right:4px solid ${m.isRead ? '#cbd5e1' : '#3b82f6'};
                              text-decoration:none;color:#1e293b;">
                        <div style="font-weight:${m.isRead ? '500' : '700'};font-size:0.85rem;margin-bottom:3px;">
                            ${m.subject || '(بدون عنوان)'}
                        </div>
                        <div style="font-size:0.72rem;color:#64748b;">
                            من: ${from} • ${date}
                        </div>
                    </a>
                `;
            });
            html += `</div>`;
        }

        if (events.length > 0) {
            html += `<div style="border-top:1px solid #e2e8f0;padding-top:10px;margin-top:6px;text-align:right;">
                <div style="font-weight:700;color:#1e293b;font-size:0.85rem;margin-bottom:6px;">
                    <i class="ti ti-calendar-event"></i> المواعيد القادمة
                </div>`;
            events.slice(0, 5).forEach(ev => {
                const start = new Date(ev.start.dateTime).toLocaleString('ar-SA', {
                    day: 'numeric', month: 'short', hour: '2-digit', minute: '2-digit'
                });
                html += `
                    <a href="${ev.webLink}" target="_blank" rel="noopener noreferrer"
                       style="display:block;padding:8px 10px;margin-bottom:5px;border-radius:8px;
                              background:#fef9f3;border-right:3px solid #f59e0b;
                              text-decoration:none;color:#78350f;font-size:0.78rem;">
                        <strong>${ev.subject || '(بدون عنوان)'}</strong><br>
                        ${start} ${ev.location?.displayName ? '• ' + ev.location.displayName : ''}
                    </a>
                `;
            });
            html += `</div>`;
        }

        html += `
            <button class="close-modal" id="closeNotificationModalBtn">إغلاق</button>
            <button class="close-modal" id="graphLogoutBtn"
                    style="background:#64748b;margin-top:8px;font-size:0.78rem;">
                <i class="ti ti-logout"></i> تسجيل الخروج
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
                const badge = document.querySelector('.notification-count');
                if (badge) badge.style.display = 'none';
            });

    } catch (err) {
        console.error('❌ Graph notifications error:', err);
        modalContent.innerHTML = `
            <div class="modal-logo"><img src="https://i.postimg.cc/bJ4RsLjj/email-(2).png" alt="إشعارات"></div>
            <div class="support-text">تعذّر تحميل الإشعارات.<br><small style="color:#64748b;font-size:0.75rem;">${err.message || ''}</small></div>
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
    if (btn._graphBound) {
        console.log('ℹ️ Already bound');
        return;
    }
    btn._graphBound = true;
    btn.addEventListener('click', function() {
        console.log('🔔 Bell clicked!');
        const modal = document.getElementById('notificationModal');
        if (modal) modal.style.display = 'flex';
        window.loadRealNotifications();
    });
    console.log('✅ Notification button bound');
}

/* ===== التهيئة الأولية عند تحميل الصفحة ===== */
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
