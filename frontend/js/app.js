/* Single-file roster app: full CRUD tabs, smart generator, printing/export helpers */
    const defaultRanks = ["ملازم","ملازم أول","نقيب","رائد","مقدم","عقيد","عميد","لواء"];
    let ranks = [];
    const weekdays = ["الأحد","الاثنين","الثلاثاء","الأربعاء","الخميس","الجمعة","السبت"];
     const THEMES = {
      classic: { name:'كلاسيكي', brand:'#1e40af', brandDark:'#0f172a', bg:'#f0f9ff', text:'#111' },
      night: { name:'ليلي', brand:'#0ea5e9', brandDark:'#0b172a', bg:'#0f172a', text:'#e5e7eb' },
      desert: { name:'رملي', brand:'#b45309', brandDark:'#78350f', bg:'#fef3c7', text:'#1f2937' },
      mint: { name:'نعناعي', brand:'#0d9488', brandDark:'#0f766e', bg:'#ecfeff', text:'#0f172a' }
    };
    const fontOptions = [
      {label:'Arial', value:'Arial'},
      {label:'Tahoma', value:'Tahoma'},
      {label:'Noto Sans Arabic', value:'"Noto Sans Arabic"'},
      {label:'Cairo', value:'Cairo'},
      {label:'Tajawal', value:'Tajawal'},
      {label:'Almarai', value:'Almarai'},
      {label:'Amiri', value:'Amiri'},
      {label:'Dubai', value:'Dubai'},
      {label:'Fanan', value:'Fanan'},
      {label:'Sultan bold', value:'"Sultan bold"'},
      {label:'PT Bold Heading', value:'"PT Bold Heading"'},
      {label:'خط مخصص', value:'AppCustomFont'}
    ];

    let officers = [];
    let officerLimits = {};
    const officerLimitFilters = { hiddenRanks: [], hideInternal: false, hideExternal: false, rankLimitExcluded: [] };
    let departments = [];
    let jobTitles = [];
    let duties = [];
    let roster = {};
    let archivedRoster = {};
    let exceptions = [];
    let activityLog = [];
    let SELECTED_OFFICER_ID = null;
    let supportRequests = [];
    const FIRST_RUN_KEY = 'rosterFirstRunDone';
    const OFFICER_RANDOMNESS = {};
    const dutyCollapseState = {};
    const dutySwapState = {};
    const deptCollapseState = {};
    let ADMIN_RECOVERY_CONTEXT = null;
    function parseBadgeSegment(val){
      const n = parseInt(val, 10);
      return isNaN(n) ? Number.MAX_SAFE_INTEGER : n;
    }
    function splitBadgeParts(badge){
      const parts = String(badge || '').split('/');
      return [parts[0] || '', parts[1] || '', parts[2] || ''];
    }
    function buildBadgeValue(part1, part2, part3){
      const p1 = (part1 || '').trim();
      const p2 = (part2 || '').trim();
      const p3 = (part3 || '').trim();
      if(!p1 && !p2 && !p3) return '';
      return `${p1}/${p2}${p3 ? `/${p3}` : ''}`;
    }
    function getBadgeSortKey(badge){
      const [first, second, third] = splitBadgeParts(badge);
      return [
        parseBadgeSegment(second),
        parseBadgeSegment(first),
        parseBadgeSegment(third)
      ];
    }
    function compareOfficerFairness(a, b, totalsMap){
      const totalA = totalsMap?.[a.id]?.total || totalsMap?.[a.id] || 0;
      const totalB = totalsMap?.[b.id]?.total || totalsMap?.[b.id] || 0;
      if(totalA !== totalB) return totalA - totalB;
      const rankDiff = ranks.indexOf(a.rank) - ranks.indexOf(b.rank);
      if(rankDiff !== 0) return rankDiff;
      const badgeA = getBadgeSortKey(a.badge);
      const badgeB = getBadgeSortKey(b.badge);
      for(let i=0; i<badgeA.length; i++){
        const diff = badgeA[i] - badgeB[i];
        if(diff !== 0) return diff;
      }
      return (a.name||'').localeCompare(b.name||'');
    }
    function getWeekIndexForDate(dateObj){
      if(!dateObj) return 0;
      const firstDay = new Date(dateObj.getFullYear(), dateObj.getMonth(), 1).getDay();
      const offset = (firstDay - 6 + 7) % 7;
      return Math.floor(((dateObj.getDate() - 1) + offset) / 7);
    }
    function ruleMatchesWeekday(rule, weekday, dateObj=null){
      const days = Array.isArray(rule.weekdays) && rule.weekdays.length ? rule.weekdays.map(Number) : null;
      if(days && rule.weekdayMode === 'alternate' && days.length > 1 && dateObj){
        const weekIndex = getWeekIndexForDate(dateObj);
        return Number(weekday) === Number(days[weekIndex % days.length]);
      }
      if(days) return days.includes(Number(weekday));
      if(rule.weekday == null) return true;
      return +rule.weekday === Number(weekday);
    }
    function normalizeOfficerKey(name, rank, badge){
      return [name || '', rank || '', badge || ''].map(v=>String(v).trim().toLowerCase()).join('|');
    }
    function isDuplicateOfficer(name, rank, badge, ignoreId=null){
      const key = normalizeOfficerKey(name, rank, badge);
      return officers.some(o => o.id !== ignoreId && normalizeOfficerKey(o.name, o.rank, o.badge) === key);
    }
   
    const defaultSettings = {
      appName: "تطبيق خدمات الإدارة العامة للمساعدات الفنية",
      appSubtitle: "",
      headerHtml: "",
      headerAlign: "center",
      logoData: "data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHZpZXdCb3g9IjAgMCA1MTIgNTEyIj4KICA8ZGVmcz4KICAgIDxsaW5lYXJHcmFkaWVudCBpZD0iYmciIHgxPSIwJSIgeDI9IjEwMCUiPgogICAgICA8c3RvcCBvZmZzZXQ9IjAlIiBzdG9wLWNvbG9yPSIjZjdmMmRlIi8+CiAgICAgIDxzdG9wIG9mZnNldD0iNTAlIiBzdG9wLWNvbG9yPSIjZjdmMmRlIi8+CiAgICAgIDxzdG9wIG9mZnNldD0iNTAlIiBzdG9wLWNvbG9yPSIjZTRkN2JhIi8+CiAgICAgIDxzdG9wIG9mZnNldD0iMTAwJSIgc3RvcC1jb2xvcj0iI2U0ZDdiYSIvPgogICAgPC9saW5lYXJHcmFkaWVudD4KICA8L2RlZnM+CiAgPGNpcmNsZSBjeD0iMjU2IiBjeT0iMjU2IiByPSIyNDAiIGZpbGw9InVybCgjYmcpIiBzdHJva2U9IiMxZTQwYWIiIHN0cm9rZS13aWR0aD0iOCIgc3Ryb2tlLWxpbmVqb2luPSJyb3VuZCIgLz4KICA8cG9seWxpbmUgcG9pbnRzPSIxNTAsMTIwIDE5MCw0MjAgMjIwLDEyMCAyNjAsNDIwIiBmaWxsPSJub25lIiBzdHJva2U9IiMxZTQwYWIiIHN0cm9rZS13aWR0aD0iMTgiIHN0cm9rZS1saW5lY2FwPSJyb3VuZCIgLz4KICA8cG9seWdvbiBwb2ludHM9IjI2MCwxMDAgMzcwLDIyMCAzMTAsMjMwIDQxMCw0MjAgMjkwLDMyMCAyNTAsNDIwIDI0MCwyMzAgMTgwLDIyMCIgZmlsbD0iIzI3MzI2YiIgLz4KICA8dGV4dCB4PSIyNTYiIHk9IjU1IiB0ZXh0LWFuY2hvcj0ibWlkZGxlIiBmb250LXNpemU9IjI4IiBmb250LWZhbWlseT0iQXJpYWwsIHNhbnMtc2VyaWYiIGZpbGw9IiMxZTQwYWIiPtix2KfZhNi52KjZiNin2Kog2YPYqNio2YrYp9mGPC90ZXh0PgogIDx0ZXh0IHg9IjI1NiIgeT0iNDc1IiB0ZXh0LWFuY2hvcj0ibWlkZGxlIiBmb250LXNpemU9IjIwIiBmb250LWZhbWlseT0iQXJpYWwsIHNhbnMtc2VyaWYiIGZpbGw9IiMxZTQwYWIiPtmF2YTZitmHINin2YTYo9iz2LnZiNmKINmF2YbYp9mI2Lkg2KjYp9mG2Ycg2KfZhNmF2YrYp9mE2YrZg9ipPC90ZXh0Pgo8L3N2Zz4=",
      footerHtml: "<div style='font-size:12px;color:#333;'>© حقوق محفوظة</div>",
      signatures: [{name:"",title:"",place:""},{name:"",title:"",place:""},{name:"",title:"",place:""}],
      authEnabled: false,
      users: [],
      printFontSize: 9,
      reportFontFamily: 'Fanan',
      appFontFamily: 'Arial',
      formFontFamily: 'Arial',
      tabFontFamily: 'Arial',
      includeHeaderOnPrint: false,
      includeFooterOnPrint: true,
      idleLogoutMinutes: 0,
      themeKey: 'classic',
      loginLayout: 'stacked',
      apiBaseUrl: '',
      apiSyncEnabled: true,
       deptGroups: [
        { key: 'internal', name: 'الإدارة العامة للمساعدات الفنية - ديوان الإدارة' },
        { key: 'external', name: 'الإدارة العامة للمساعدات الفنية- المناطق الجغرافية' },
        { key: 'transfer', name: 'جهات النقل (إدارات عامة ومديريات الأمن)' }
      ]
    };
    const REPORT_MAX_FONT_SIZE = 10;
    let SETTINGS = Object.assign({}, defaultSettings);
     let CURRENT_USER = null;
    let ACTIVE_SESSIONS = [];
      const ROLE_TABS = {
    admin: ['roster','officers','ranks','departments','transfer-departments','jobtitles','duties','limits','exceptions','report','stats','admin','settings','account','about'],
      editor: ['roster','report','stats','officers','account','about'], // no org/rules edits
      user: ['roster','account','about'],
      viewer: ['roster','account','about']
    };
    const TAB_OPTIONS = [
      {id:'roster', label:'الخدمات'},
      {id:'officers', label:'الضباط'},
      {id:'ranks', label:'الرتب'},
      {id:'departments', label:'الأقسام'},
      {id:'transfer-departments', label:'جهات خارجية'},
      {id:'jobtitles', label:'المسميات الوظيفية'},
      {id:'duties', label:'أنواع الخدمات'},
      {id:'limits', label:'حدود الضباط'},
      {id:'exceptions', label:'القواعد'},
      {id:'report', label:'التقرير'},
      {id:'stats', label:'إحصاءات/أرشيف'},
      {id:'account', label:'الملف الشخصي'},
      {id:'about', label:'حول التطبيق'}
    ];

    /* ========= Persistence ========= */
    function safeLoadFromStorage(key, fallback){
      try {
        const raw = localStorage.getItem(key);
        if(raw === null || raw === undefined) return fallback;
        const parsed = JSON.parse(raw);
        return parsed === null || parsed === undefined ? fallback : parsed;
      } catch(_) {
        return fallback;
      }
    }
    function shouldShowFirstRunPrompt(){
      if(localStorage.getItem(FIRST_RUN_KEY) === 'true') return false;
      const keys = ['officers','departments','jobTitles','duties','roster','exceptions','ranks','settings'];
      const hasAnyStored = keys.some(key => localStorage.getItem(key) !== null);
      if(!hasAnyStored) return true;
      const storedOfficers = safeLoadFromStorage('officers', []);
      const storedRoster = safeLoadFromStorage('roster', {});
      const storedSettings = safeLoadFromStorage('settings', {});
      const storedUsers = Array.isArray(storedSettings.users) ? storedSettings.users : [];
      const hasRoster = storedRoster && typeof storedRoster === 'object' && Object.keys(storedRoster).length > 0;
      return !(Array.isArray(storedOfficers) && storedOfficers.length) && !hasRoster && !storedUsers.length;
    }
    function completeFirstRunSetup(){
      localStorage.setItem(FIRST_RUN_KEY, 'true');
      saveAll();
      showToast('تم إعداد النظام للبدء.', 'success');
      renderTab('roster');
    }
    function handleFirstRunRestore(file){
      if(!file) return;
      importFullBackup(file, {
        onComplete: () => {
          localStorage.setItem(FIRST_RUN_KEY, 'true');
          renderTab('roster');
        }
      });
    }
    function getDefaultDuties(){
      return [
        {id:1,name:"إدارية",printLabel:"إدارية",color:"danger",duration:8,restHours:24,allowedRanks:[],signingJobTitleIds:[]},{id:2,name:"ميدانية",printLabel:"ميدانية",color:"success",duration:12,restHours:36,allowedRanks:[],signingJobTitleIds:[]},
        {id:3,name:"احتياط",printLabel:"احتياط",color:"warning",duration:24,restHours:48,allowedRanks:[],signingJobTitleIds:[]},
        {id:4,name:"ليلية",printLabel:"ليلية",color:"dark",duration:10,restHours:38,allowedRanks:[],signingJobTitleIds:[]}
      ];
    }

   function getDefaultDepartments(){
      return [
        {id:1,name:"الإدارة العامة للمساعدات الفنية",parentId:null,headId:null,upperTitle:'',upperOfficerId:null,groupKey:'internal'},
        {id:2,name:"الإدارة العامة ",parentId:null,headId:null,upperTitle:'',upperOfficerId:null,groupKey:'external'},
        {id:3,name:"مديرية أمن",parentId:null,headId:null,upperTitle:'',upperOfficerId:null,groupKey:'external'}
      ];
    }

    const REMOTE_SAVE_DELAY_MS = 1200;
    let remoteSaveTimer = null;
    let remoteSaveInFlight = false;
    let remoteSaveQueued = false;
    let remoteLoadAttempted = false;

    function getApiBaseUrl(){
      const configured = (SETTINGS.apiBaseUrl || '').trim();
      if(configured) return configured.replace(/\/$/, '');
      if(window.location.protocol === 'http:' || window.location.protocol === 'https:'){
        return window.location.origin;
      }
      return '';
    }
    function isRemoteSyncEnabled(){
      return !!SETTINGS.apiSyncEnabled && !!getApiBaseUrl();
    }
    function buildRemotePayload(){
      return {
        officers,
        officerLimits,
        departments,
        jobTitles,
        duties,
        roster,
        archivedRoster,
        exceptions,
        ranks,
        settings: SETTINGS,
        activityLog,
        supportRequests,
        officerLimitFilters,
        meta: {
          updatedAt: new Date().toISOString(),
          updatedBy: CURRENT_USER ? CURRENT_USER.name : 'system'
        }
      };
    }
    function hasRemoteData(data){
      if(!data || typeof data !== 'object') return false;
      const keys = Object.keys(data);
      return keys.some(key=>{
        const value = data[key];
        if(Array.isArray(value)) return value.length > 0;
        if(value && typeof value === 'object') return Object.keys(value).length > 0;
        return false;
      });
    }
    function applyRemotePayload(data){
      const storageMap = {
        officers: 'officers',
        officerLimits: 'officerLimits',
        departments: 'departments',
        jobTitles: 'jobTitles',
        duties: 'duties',
        roster: 'roster',
        archivedRoster: 'archivedRoster',
        exceptions: 'exceptions',
        ranks: 'ranks',
        settings: 'settings',
        activityLog: 'activityLog',
        supportRequests: 'supportRequests',
        officerLimitFilters: 'officerLimitFilters'
      };
      Object.keys(storageMap).forEach(key=>{
        if(data[key] !== undefined){
          localStorage.setItem(storageMap[key], JSON.stringify(data[key]));
        }
      });
    }
    function requestRemoteLoad(){
      if(!isRemoteSyncEnabled() || remoteLoadAttempted) return;
      remoteLoadAttempted = true;
      fetch(`${getApiBaseUrl()}/api/data`, { headers: { 'Accept': 'application/json' } })
        .then(res => res.ok ? res.json() : null)
        .then(payload=>{
          if(!payload) return;
          const data = payload.data || payload;
          if(!hasRemoteData(data)) return;
          applyRemotePayload(data);
          loadAll({ skipRemoteLoad: true });
          if(splashCompleted){
            renderTab('roster');
          }
        })
        .catch(err=> console.warn('Remote load failed', err));
    }
    function scheduleRemoteSave(){
      if(!isRemoteSyncEnabled()) return;
      if(remoteSaveTimer) clearTimeout(remoteSaveTimer);
      remoteSaveTimer = setTimeout(()=>{
        remoteSaveTimer = null;
        sendRemoteSave();
      }, REMOTE_SAVE_DELAY_MS);
    }
    async function sendRemoteSave(){
      if(!isRemoteSyncEnabled()) return;
      if(remoteSaveInFlight){
        remoteSaveQueued = true;
        return;
      }
      remoteSaveInFlight = true;
      try {
        await fetch(`${getApiBaseUrl()}/api/data`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ data: buildRemotePayload() })
        });
      } catch (err){
        console.warn('Remote save failed', err);
      } finally {
        remoteSaveInFlight = false;
        if(remoteSaveQueued){
          remoteSaveQueued = false;
          sendRemoteSave();
        }
      }
    }

    function loadAll(options = {}){
      try { officers = JSON.parse(localStorage.getItem('officers')) || []; } catch(_) { officers = []; }
      try { officerLimits = JSON.parse(localStorage.getItem('officerLimits')) || {}; } catch(_) { officerLimits = {}; }
      try { departments = JSON.parse(localStorage.getItem('departments')) || getDefaultDepartments(); } catch(_) { departments = getDefaultDepartments(); }
      try { jobTitles = JSON.parse(localStorage.getItem('jobTitles')) || [{id:1,name:"قائد الوحدة",parentId:null,isUpper:true},{id:2,name:"رئيس قسم الشئون الإدارية",parentId:null,isUpper:false}]; } catch(_) { jobTitles = [{id:1,name:"قائد الوحدة",parentId:null,isUpper:true},{id:2,name:"رئيس قسم الشئون الإدارية",parentId:null,isUpper:false}]; }
      try { duties = JSON.parse(localStorage.getItem('duties')) || getDefaultDuties(); } catch(_) { duties = getDefaultDuties(); }
      try { roster = JSON.parse(localStorage.getItem('roster')) || {}; } catch(_) { roster = {}; }
      try { archivedRoster = JSON.parse(localStorage.getItem('archivedRoster')) || {}; } catch(_) { archivedRoster = {}; }
      try { exceptions = JSON.parse(localStorage.getItem('exceptions')) || []; } catch(_) { exceptions = []; }
      try { activityLog = JSON.parse(localStorage.getItem('activityLog')) || []; } catch(_) { activityLog = []; }
      try { ranks = JSON.parse(localStorage.getItem('ranks')) || defaultRanks.slice(); } catch(_) { ranks = defaultRanks.slice(); }
      if(!Array.isArray(ranks) || !ranks.length) ranks = defaultRanks.slice();
      try { Object.assign(officerLimitFilters, JSON.parse(localStorage.getItem('officerLimitFilters')) || {}); } catch(_) {}
      try { SETTINGS = Object.assign(SETTINGS, JSON.parse(localStorage.getItem('settings')) || {}); } catch(_) {}
      if(!Array.isArray(SETTINGS.users)) SETTINGS.users = [];
      SETTINGS.users = (SETTINGS.users||[]).map(u=>({
        name: u.name||'',
        pwd: u.pwd||'',
        role: u.role||'user',
        officerId: u.officerId ?? null,
        phone: u.phone||'',
        note: u.note||'',
        fullName: u.fullName || u.name || '',
        email: u.email || '',
        tabPrivileges: Array.isArray(u.tabPrivileges) ? u.tabPrivileges : [],
        createdAt: u.createdAt || new Date().toISOString(),
        lastLoginAt: u.lastLoginAt || null,
        mustChangePassword: u.mustChangePassword === undefined ? !(u.lastLoginAt) : !!u.mustChangePassword
      }));
      if(!Array.isArray(SETTINGS.deptGroups) || !SETTINGS.deptGroups.length) SETTINGS.deptGroups = defaultSettings.deptGroups.slice();
      if(!SETTINGS.themeKey || !THEMES[SETTINGS.themeKey]) SETTINGS.themeKey = defaultSettings.themeKey;
      if(!SETTINGS.loginLayout) SETTINGS.loginLayout = defaultSettings.loginLayout;
      if(!SETTINGS.reportFontFamily) SETTINGS.reportFontFamily = defaultSettings.reportFontFamily;
      if(!SETTINGS.appFontFamily) SETTINGS.appFontFamily = defaultSettings.appFontFamily;
      if(!SETTINGS.formFontFamily) SETTINGS.formFontFamily = defaultSettings.formFontFamily;
      if(!SETTINGS.tabFontFamily) SETTINGS.tabFontFamily = defaultSettings.tabFontFamily;
      if(!SETTINGS.reportFontFamily) SETTINGS.reportFontFamily = defaultSettings.reportFontFamily;
      if(SETTINGS.apiSyncEnabled === undefined) SETTINGS.apiSyncEnabled = defaultSettings.apiSyncEnabled;
      if(SETTINGS.apiBaseUrl === undefined) SETTINGS.apiBaseUrl = defaultSettings.apiBaseUrl;
      try { ACTIVE_SESSIONS = JSON.parse(sessionStorage.getItem('activeSessions')) || []; } catch(_) { ACTIVE_SESSIONS = []; }
      try { CURRENT_USER = JSON.parse(sessionStorage.getItem('currentUser')) || null; if(CURRENT_USER && CURRENT_USER.officerId===undefined) CURRENT_USER.officerId=null; } catch(_) { CURRENT_USER = null; }
      try { supportRequests = JSON.parse(localStorage.getItem('supportRequests')) || []; } catch(_) { supportRequests = []; }
      document.getElementById('appNameTitle').textContent = SETTINGS.appName;
      document.getElementById('appSubtitle').textContent = SETTINGS.appSubtitle || '';
      const footer = document.getElementById('appFooterText');
      if(footer){
        const year = new Date().getFullYear();
        footer.textContent = `© ${year} ${SETTINGS.appName} — جميع الحقوق محفوظة`;
      }
      updateLoginLabel();
      initSplashScreen();
      if(CURRENT_USER){
        const rec = (SETTINGS.users||[]).find(u=>u.name===CURRENT_USER.name);
        if(rec){
          CURRENT_USER.mustChangePassword = !!rec.mustChangePassword;
          CURRENT_USER.fullName = rec.fullName || rec.name || CURRENT_USER.fullName || CURRENT_USER.name || '';
          CURRENT_USER.tabPrivileges = Array.isArray(rec.tabPrivileges) ? rec.tabPrivileges : [];
          sessionStorage.setItem('currentUser', JSON.stringify(CURRENT_USER));
        }
      }
         officers = officers.map(o=>{
          const normalized = Object.assign({status:'active',transferNote:'',phone:'',gender:'',birth:'',address:'',emgName:'',emgPhone:'',admissionType:'fresh',hiringDate:'',transferFromDeptId:null,transferToDeptId:null,transferDate:'',transferToDate:'', archivedAt:null}, o);
        if(!normalized.transferDate && normalized.transferFromDate) normalized.transferDate = normalized.transferFromDate;
        return normalized;
      });
      departments = departments.map(d=>Object.assign({groupKey:'internal'}, d));
      applyNavPermissions();
      applyTheme();
      applyCustomFont();
      applyFontSettings();
      updateNavLogo();
      cleanupOfficerLimits();
      autoArchiveOldMonths();
      resetIdleTimer();
      if(SETTINGS.apiSyncEnabled === undefined) SETTINGS.apiSyncEnabled = defaultSettings.apiSyncEnabled;
      if(SETTINGS.apiBaseUrl === undefined) SETTINGS.apiBaseUrl = defaultSettings.apiBaseUrl;
    }

    function saveAll(options = {}){
      localStorage.setItem('officers', JSON.stringify(officers));
      localStorage.setItem('officerLimits', JSON.stringify(officerLimits));
      localStorage.setItem('departments', JSON.stringify(departments));
      localStorage.setItem('jobTitles', JSON.stringify(jobTitles));
      localStorage.setItem('duties', JSON.stringify(duties));
      localStorage.setItem('ranks', JSON.stringify(ranks));
      localStorage.setItem('roster', JSON.stringify(roster));
      localStorage.setItem('archivedRoster', JSON.stringify(archivedRoster));
      localStorage.setItem('exceptions', JSON.stringify(exceptions));
      localStorage.setItem('supportRequests', JSON.stringify(supportRequests));
      localStorage.setItem('officerLimitFilters', JSON.stringify(officerLimitFilters));
      activityLog = activityLog.slice(-500);
      localStorage.setItem('activityLog', JSON.stringify(activityLog));
      localStorage.setItem('settings', JSON.stringify(SETTINGS));
      localStorage.removeItem('activeSessions');
      sessionStorage.setItem('activeSessions', JSON.stringify(ACTIVE_SESSIONS));
      sessionStorage.setItem('currentUser', JSON.stringify(CURRENT_USER));
      if(!options.skipRemote){
        scheduleRemoteSave();
      }
     }


    let idleTimer = null;
    function resetIdleTimer(){
      if(idleTimer) clearTimeout(idleTimer);
      if(!SETTINGS.authEnabled || !CURRENT_USER || !SETTINGS.idleLogoutMinutes) return;
      idleTimer = setTimeout(()=>{ showToast('تم تسجيل الخروج بسبب الخمول','warning'); signOut(true); }, SETTINGS.idleLogoutMinutes*60*1000);
    }
    function clearIdleTimer(){ if(idleTimer) clearTimeout(idleTimer); idleTimer=null; }

    /* ========= Helpers ========= */
    function escapeHtml(s){ if(s==null) return ''; return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/"/g,'&quot;'); }
    function csvEscape(v){ if(v==null) return ''; v = String(v); if(v.includes(',') || v.includes('"')) return `"${v.replace(/"/g,'""')}"`; return v; }
    function downloadJSONFile(name,content){ const blob=new Blob([content],{type:'application/json;charset=utf-8'}); const a=document.createElement('a'); a.href=URL.createObjectURL(blob); a.setAttribute('download',name); document.body.appendChild(a); a.click(); a.remove(); }
    function downloadCSVFile(name,content){ const blob=new Blob([content],{type:'text/csv;charset=utf-8;'}); const a=document.createElement('a'); a.href=URL.createObjectURL(blob); a.setAttribute('download',name); document.body.appendChild(a); a.click(); a.remove(); }
    function normalizeLimit(val){ if(val===0 || val==='' || val==null) return null; return +val; }
    function ensureOfficerLimit(officerId){ if(!officerLimits[officerId]) officerLimits[officerId] = { duties:{}, total:null }; if(!officerLimits[officerId].duties) officerLimits[officerId].duties = {}; }
    function cleanupOfficerLimits(){
      const valid = new Set(officers.map(o=>String(o.id)));
      const cleaned = {};
      Object.keys(officerLimits||{}).forEach(id=>{
        if(!valid.has(String(id))) return;
        const rec = officerLimits[id] || {};
        const dutiesLimits = {};
        Object.keys(rec.duties||{}).forEach(dId=>{
          const v = normalizeLimit(rec.duties[dId]);
          if(v != null) dutiesLimits[dId] = v;
        });
        cleaned[id] = { duties: dutiesLimits, total: normalizeLimit(rec.total) };
      });
      officerLimits = cleaned;
    }
    function getOfficerDutyLimit(officerId, dutyId){ const rec = officerLimits[officerId]; if(!rec || rec.duties==null) return null; const v = rec.duties[dutyId]; return v==null ? null : +v; }
    function getOfficerTotalLimit(officerId){ const rec = officerLimits[officerId]; if(!rec) return null; return rec.total==null ? null : +rec.total; }
    function setOfficerDutyLimit(officerId, dutyId, val){ ensureOfficerLimit(officerId); officerLimits[officerId].duties[dutyId] = normalizeLimit(val); saveAll(); renderTab('limits'); }
    function setOfficerTotalLimit(officerId, val){ ensureOfficerLimit(officerId); officerLimits[officerId].total = normalizeLimit(val); saveAll(); renderTab('limits'); }
    function currentRole(){ return SETTINGS.authEnabled ? (CURRENT_USER?.role || 'user') : 'admin'; }
    function isPasswordChangeRequired(){ return SETTINGS.authEnabled && CURRENT_USER?.mustChangePassword; }
    function getRoleTabs(role){ return ROLE_TABS[role] || ['roster']; }
    function getEffectiveTabsForUser(user){
      if(!user) return getRoleTabs(currentRole());
      const roleTabs = getRoleTabs(user.role);
      if((user.role === 'user' || user.role === 'editor') && Array.isArray(user.tabPrivileges) && user.tabPrivileges.length){
        return user.tabPrivileges;
      }
      return roleTabs;
    }
    function getAllowedTabsForCurrentUser(){
      if(!SETTINGS.authEnabled) return ROLE_TABS.admin;
      return getEffectiveTabsForUser(currentUserRecord());
    }
    function tabAllowed(tab){ const allowed = getAllowedTabsForCurrentUser() || ['roster']; return allowed.includes(tab); }
    function applyNavPermissions(){
      const forcedTabs = isPasswordChangeRequired() ? ['account'] : null;
      const allowed = forcedTabs || getAllowedTabsForCurrentUser() || [];
      document.querySelectorAll('#mainTabs .nav-link').forEach(btn=>{
        const can = !SETTINGS.authEnabled || allowed.includes(btn.dataset.tab);
        btn.classList.toggle('disabled', !can);
      });
    }
    function applyTheme(){
      const theme = THEMES[SETTINGS.themeKey] || THEMES.classic;
      const root = document.documentElement;
      root.style.setProperty('--brand', theme.brand);
      root.style.setProperty('--brand-dark', theme.brandDark);
      root.style.setProperty('--bg', theme.bg);
      root.style.setProperty('--text', theme.text);
      document.body.style.background = `linear-gradient(135deg, ${theme.bg}, ${theme.brand}10)`;
    }
    function applyCustomFont(){
      if(!SETTINGS.customFontData) return;
      let styleEl = document.getElementById('custom_font_style');
      if(!styleEl){ styleEl = document.createElement('style'); styleEl.id='custom_font_style'; document.head.appendChild(styleEl); }
      styleEl.textContent = `
        @font-face { font-family: "AppCustomFont"; src: url(${SETTINGS.customFontData}); font-weight: 300 900; }
      `;
    }
    function applyFontSettings(){
      const root = document.documentElement;
      const fallback = '"Arial","Noto Sans Arabic","Tahoma",sans-serif';
      root.style.setProperty('--app-font', `${SETTINGS.appFontFamily || 'Arial'}, ${fallback}`);
      root.style.setProperty('--form-font', `${SETTINGS.formFontFamily || 'Arial'}, ${fallback}`);
      root.style.setProperty('--tab-font', `${SETTINGS.tabFontFamily || 'Arial'}, ${fallback}`);
    }
    function updateNavLogo(){
      const logo = document.getElementById('appLogoImg');
      if(!logo) return;
      if(SETTINGS.logoData){
        logo.src = SETTINGS.logoData;
        logo.style.display = 'inline-block';
      } else {
        logo.removeAttribute('src');
        logo.style.display = 'none';
      }
    }
    function hideSplashScreen(){
      const splash = document.getElementById('splashScreen');
      if(!splash || splash.classList.contains('hidden')) return;
      splash.classList.add('hidden');
      setTimeout(()=> splash.remove(), 1200);
    }
    function initSplashScreen(){
      const splash = document.getElementById('splashScreen');
      if(!splash) return;
      const logo = document.getElementById('splashLogoImg');
      if(logo){
        const splashLogo = SETTINGS.logoData || defaultSettings.logoData;
        if(splashLogo){ logo.src = splashLogo; }
      }
      const title = splash.querySelector('.splash-title');
      if(title) title.textContent = `مرحباً بك في ${SETTINGS.appName}`;
    }
    function updateSplashProgress(value){
      const bar = document.getElementById('splashProgressBar');
      const label = document.getElementById('splashProgressLabel');
      const safeValue = Math.min(100, Math.max(0, Math.round(value)));
      if(bar) bar.style.width = `${safeValue}%`;
      if(label) label.textContent = `جارٍ التحضير... ${safeValue}%`;
    }
    let splashProgressTimer = null;
    let splashCompleted = false;
    function startSplashProgress(){
      const splash = document.getElementById('splashScreen');
      if(!splash){
        renderTab('roster');
        return;
      }
      let progress = 0;
      updateSplashProgress(progress);
      const steps = [8, 18, 35, 52, 68, 82, 92, 100];
      let idx = 0;
      splashProgressTimer = setInterval(()=>{
        progress = steps[idx] ?? 100;
        updateSplashProgress(progress);
        idx++;
        if(progress >= 100){
          clearInterval(splashProgressTimer);
          splashProgressTimer = null;
          if(!splashCompleted){
            splashCompleted = true;
            setTimeout(()=>{
              renderTab('roster');
              hideSplashScreen();
            }, 300);
          }
        }
      }, 260);
    }
    function logActivity(action, meta = {}){
      const actorName = meta.actor || (CURRENT_USER ? CURRENT_USER.name : 'مجهول');
      const actorRole = meta.actorRole || (CURRENT_USER ? CURRENT_USER.role : 'guest');
      const entry = {
        action,
        meta,
        user: actorName,
        role: actorRole,
        ts: new Date().toISOString()
      };
      activityLog.push(entry);
      activityLog = activityLog.slice(-500);
      try { localStorage.setItem('activityLog', JSON.stringify(activityLog)); } catch(_) {}
    }
    function passwordStrengthScore(pwd){
      let score = 0;
      if(!pwd) return score;
      if(pwd.length >= 8) score++;
      if(/[A-Z؀-ۿ]/.test(pwd) && /[a-z]/.test(pwd)) score++; // mix of cases/letters
      if(/[0-9]/.test(pwd)) score++;
      if(/[^A-Za-z0-9]/.test(pwd)) score++;
      if(pwd.length >= 12) score++;
      return Math.min(score, 4);
    }
    function updateStrengthBar(pwd, holderId){
      const el = document.getElementById(holderId);
      if(!el) return;
      const bar = el.querySelector('.bar');
      if(!bar) return;
      const score = passwordStrengthScore(pwd);
      const pct = (score/4)*100;
      const colors = ['#ef4444','#f97316','#facc15','#22c55e','#15803d'];
      bar.style.width = pct+'%';
      bar.style.background = colors[score];
    }
    function formatActivityDetails(entry){
      if(!entry || !entry.meta) return '';
      const parts = [];
      if(entry.meta.month) parts.push('شهر: '+entry.meta.month);
      if(entry.meta.dutyName) parts.push('خدمة: '+entry.meta.dutyName);
      if(entry.meta.day) parts.push('يوم: '+entry.meta.day);
      if(entry.meta.note) parts.push(entry.meta.note);
      if(entry.meta.target) parts.push(entry.meta.target);
      return parts.join(' • ');
    }
    
    
    function getDeptGroup(deptId){
      const d = departments.find(dd=>dd.id===deptId);
      return d?.groupKey || 'internal';
    }
    function canEditRoster(){ const r=currentRole(); return !SETTINGS.authEnabled || r==='admin' || r==='editor'; }
    function isViewerUser(){ return SETTINGS.authEnabled && currentRole()==='viewer'; }
    function isViewOnlyUser(){ return SETTINGS.authEnabled && (currentRole()==='user' || currentRole()==='viewer'); }
    // Offline fallbacks for missing libs
    // Lightweight html2canvas alternative using SVG foreignObject to avoid external dependencies
    if (typeof window.html2canvas === 'undefined') {
      window.html2canvas = function(element, opts={}){
        return new Promise((resolve, reject) => {
          try{
            const scale = opts.scale || 1;
            const rect = element.getBoundingClientRect();
            const width = Math.max(element.scrollWidth || rect.width, rect.width);
            const height = Math.max(element.scrollHeight || rect.height, rect.height);
            const styles = Array.from(document.querySelectorAll('style')).map(s=>s.outerHTML).join('\n');
            const serializer = new XMLSerializer();
            const clonedHtml = serializer.serializeToString(element.cloneNode(true));
            const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}"><foreignObject width="100%" height="100%"><div xmlns="http://www.w3.org/1999/xhtml">${styles}${clonedHtml}</div></foreignObject></svg>`;
            const img = new Image();
            img.onload = ()=>{
              const canvas = document.createElement('canvas');
              canvas.width = width * scale;
              canvas.height = height * scale;
              const ctx = canvas.getContext('2d');
              ctx.scale(scale, scale);
              ctx.drawImage(img, 0, 0, width, height);
              resolve(canvas);
            };
            img.onerror = reject;
            img.src = 'data:image/svg+xml;charset=utf-8,' + encodeURIComponent(svg);
          }catch(err){ reject(err); }
        });
      };
    }
    function autoArchiveOldMonths(){
      const cur = new Date().toISOString().slice(0,7);
      Object.keys(roster||{}).forEach(m=>{
        if(m < cur){
          archivedRoster[m] = roster[m];
          delete roster[m];
        }
      });
      saveAll();
    }
    function getLastAssignmentInfo(monthStr, officerId, dayCutoff, tempRoster=null){
      let bestDay=null, restDays=0;
      const consider = (dId, day) => {
        if(day>=dayCutoff) return;
        if(bestDay===null || day>bestDay){
          const dutyObj = duties.find(dd=>dd.id==dId);
          bestDay = day;
          restDays = Math.ceil(((dutyObj?.restHours)||0)/24);
        }
      };
      const monthRoster = roster[monthStr] || {};
      Object.keys(monthRoster).forEach(dId=>{
        (monthRoster[dId]||[]).forEach(r=>{ if(r.officerId===officerId) consider(+dId, r.day); });
      });
      if(tempRoster){
        Object.keys(tempRoster).forEach(dId=>{
          (tempRoster[dId]||[]).forEach((r,idx)=>{ const d = r.day || idx+1; if(r.officerId===officerId) consider(+dId, d); });
        });
      }
      return bestDay!=null ? { day: bestDay, restDays } : null;
    }

    /* ========= Tabs ========= */
    function clearActiveTab(){ document.querySelectorAll('#mainTabs .nav-link').forEach(n=>n.classList.remove('active')); }
    function renderFirstRunPrompt(){
      const logoSize = Math.min(220, window.innerWidth * 0.4);
      const logo = SETTINGS.logoData
        ? `<img src="${SETTINGS.logoData}" class="logo logo-hero" style="width:${logoSize}px;height:${logoSize}px;border-radius:50%;box-shadow:0 12px 28px rgba(30,64,175,.2);" alt="logo">`
        : `<div class="logo logo-hero" style="width:${logoSize}px;height:${logoSize}px;border-radius:50%;background:radial-gradient(circle at 30% 30%,#fff,#c7d2fe);box-shadow:0 12px 28px rgba(30,64,175,.2);"></div>`;
      return `
        <div class="card first-run-card">
          <div class="card-header">
            <div>إعداد أولي للنظام</div>
            <div class="small-muted">البدء لأول مرة</div>
          </div>
          <div class="card-body">
            <div class="text-center mb-4">${logo}</div>
            <h3 class="text-center mb-2">مرحباً بك في ${escapeHtml(SETTINGS.appName || 'النظام')}</h3>
            <p class="text-center text-muted mb-4">لم نعثر على بيانات محفوظة بعد. اختر كيفية البدء.</p>
            <div class="first-run-actions">
              <div class="first-run-option">
                <h4>ابدأ تشغيل النظام لأول مرة</h4>
                <p>سننشئ الهيكل الأساسي والبيانات الافتراضية لتبدأ بإضافة الضباط والخدمات مباشرة.</p>
                <button class="btn btn-success" onclick="completeFirstRunSetup()">بدء التشغيل</button>
              </div>
              <div class="first-run-option">
                <h4>استعادة البيانات من ملف</h4>
                <p>قم بتحميل نسخة احتياطية محفوظة مسبقاً لاستعادة جميع الأقسام والضباط والإعدادات.</p>
                <input type="file" class="form-control" accept="application/json" onchange="handleFirstRunRestore(this.files[0])">
                <button class="btn btn-outline-primary" onclick="this.previousElementSibling?.click()">اختيار ملف النسخة الاحتياطية</button>
              </div>
            </div>
          </div>
        </div>`;
    }
    function renderTab(tab){
      if(shouldShowFirstRunPrompt()){
        clearActiveTab();
        document.getElementById('mainContent').innerHTML = renderFirstRunPrompt();
        return;
      }
      if(SETTINGS.authEnabled && !CURRENT_USER){
        document.getElementById('mainContent').innerHTML = renderLandingView();
        return;
      }
      if(SETTINGS.authEnabled && isPasswordChangeRequired() && tab !== 'account') tab = 'account';
      if(SETTINGS.authEnabled && !tabAllowed(tab)) tab = 'roster';
      clearActiveTab();
      const btn = Array.from(document.querySelectorAll('#mainTabs .nav-link')).find(b=>b.dataset.tab===tab);
      if(btn) btn.classList.add('active');
      const c = document.getElementById('mainContent');
      applyNavPermissions();
      switch(tab){
        case 'officers': c.innerHTML = renderOfficersTab(); break;
        case 'ranks': c.innerHTML = renderRanksTab(); break;
        case 'departments': c.innerHTML = renderDepartmentsTab(); break;
         case 'transfer-departments': c.innerHTML = renderTransferDepartmentsTab(); break;
        case 'jobtitles': c.innerHTML = renderJobTitlesTab(); break;
        case 'duties': c.innerHTML = renderDutiesTab(); break;
        case 'limits': c.innerHTML = renderOfficerLimitsTab(); break;
        case 'exceptions': c.innerHTML = renderExceptionsTab(); clearExceptionForm(); break;
        case 'report': c.innerHTML = renderReportTab(); break;
        case 'stats': c.innerHTML = renderStatsArchiveTab(); break;
        case 'account': c.innerHTML = renderAccountTab(); break;
        case 'admin':
          c.innerHTML = renderAdminTab();
          updatePrivilegeControlState();
          break;
        case 'settings': c.innerHTML = renderSettingsTab(); break;
        case 'about': c.innerHTML = renderAboutTab(); break;
        default: c.innerHTML = renderRosterTab(); break;
      }
    }

    /* ========= Ranks Tab ========= */
    function renderRanksTab(){
      if(SETTINGS.authEnabled && !CURRENT_USER) return signInInlineFormMarkup() + '<div class="alert alert-warning">سجّل الدخول لإدارة الرتب.</div>';
      const stats = computeRankStats();
      const rows = ranks.map((r,i)=>`
        <tr>
          <td class="align-middle">${i+1}</td>
          <td><input class="form-control" value="${escapeHtml(r)}" onchange="updateRankName(${i}, this.value)"></td>
          <td class="text-center">${stats[r] || 0}</td>
          <td class="text-nowrap text-end">
            <button class="btn btn-sm btn-outline-secondary me-1" ${i===0?'disabled':''} onclick="moveRank(${i}, -1)"><i class="bi bi-arrow-up"></i></button>
            <button class="btn btn-sm btn-outline-secondary me-1" ${i===ranks.length-1?'disabled':''} onclick="moveRank(${i}, 1)"><i class="bi bi-arrow-down"></i></button>
            <button class="btn btn-sm btn-danger" onclick="deleteRank(${i})">حذف</button>
          </td>
        </tr>`).join('');

      return `<div class="card">
        <div class="card-header d-flex justify-content-between align-items-center">
          <div>إدارة الرتب العسكرية</div>
          <div>
            <button class="btn btn-outline-primary btn-sm me-2" onclick="restoreDefaultRanks()">استعادة الافتراضي</button>
            <button class="btn btn-success btn-sm" onclick="addRankFromForm()">إضافة رتبة</button>
          </div>
        </div>
        <div class="card-body">
          <div class="row g-2 mb-3">
            <div class="col-md-6"><input id="new_rank_name" class="form-control" placeholder="اسم رتبة جديدة"></div>
            <div class="col-md-6 text-muted d-flex align-items-center">إجمالي الرتب: ${ranks.length}</div>
          </div>
          <div class="table-responsive">
            <table class="table table-hover align-middle">
              <thead class="table-dark"><tr><th style="width:70px">#</th><th>الاسم</th><th style="width:120px">عدد الضباط</th><th style="width:220px" class="text-end">إجراءات</th></tr></thead>
              <tbody>${rows || '<tr><td colspan="4" class="text-center text-muted">لا توجد رتب</td></tr>'}</tbody>
            </table>
          </div>
        </div>
      </div>`;
    }
    function addRankFromForm(){
      const name = (document.getElementById('new_rank_name')?.value || '').trim();
      if(!name) return showToast('أدخل اسم رتبة', 'danger');
      if(ranks.includes(name)) return showToast('الرتبة موجودة بالفعل', 'danger');
      ranks.push(name);
      saveAll();
      showToast('تمت إضافة الرتبة', 'success');
      renderTab('ranks');
    }
    function updateRankName(idx, val){
      const name = (val || '').trim();
      if(!name) return showToast('الاسم مطلوب', 'danger');
      if(ranks.some((r,i)=>r===name && i!==idx)) return showToast('اسم مكرر', 'danger');
      const old = ranks[idx];
      ranks[idx] = name;
      officers.forEach(o=>{ if(o.rank===old) o.rank = name; });
      saveAll();
      renderTab('ranks');
    }
    function deleteRank(idx){
      const name = ranks[idx];
      if(officers.some(o=>o.rank===name)) return showToast('لا يمكن حذف رتبة مستخدمة', 'danger');
      ranks.splice(idx,1);
      saveAll();
      renderTab('ranks');
    }
    function moveRank(idx, dir){
      const target = idx + dir;
      if(target < 0 || target >= ranks.length) return;
      const temp = ranks[target]; ranks[target] = ranks[idx]; ranks[idx] = temp;
      saveAll();
      renderTab('ranks');
    }
    function restoreDefaultRanks(){
      if(!confirm('سيتم استبدال الرتب الحالية بالقائمة الافتراضية. متابعة؟')) return;
      ranks = defaultRanks.slice();
      saveAll();
      renderTab('ranks');
      showToast('تمت الاستعادة', 'success');
    }

    /* ========= Auth ========= */
    function updateLoginLabel(){
      const el = document.getElementById('currentUserLabel');
      if(CURRENT_USER){
        const linked = CURRENT_USER.officerId ? officers.find(o=>o.id===CURRENT_USER.officerId) : null;
        el.textContent = `${linked? linked.name : CURRENT_USER.name} (${CURRENT_USER.role})`;
      } else {
        el.textContent = SETTINGS.authEnabled ? 'غير مسجل' : '';
      }
      const signBtn = document.getElementById('navSignIn');
      if(signBtn) signBtn.style.display = (!CURRENT_USER && SETTINGS.authEnabled) ? 'inline-block' : 'none';
      const signOutBtn = document.getElementById('navSignOut');
      if(signOutBtn) signOutBtn.style.display = (CURRENT_USER && SETTINGS.authEnabled) ? 'inline-block' : 'none';
      applyNavPermissions();
    }
    function signInInlineFormMarkup(withLogo = true, layoutMode = 'inline'){
      const logoSize = 288; // 1.8x أكبر لواجهة الدخول
      const logo = !withLogo ? '' : (SETTINGS.logoData
        ? `<div class="text-center mb-3"><img src="${SETTINGS.logoData}" class="logo logo-hero logo-anim" style="width:${logoSize}px;height:${logoSize}px;border-radius:50%;box-shadow:0 12px 32px rgba(99,102,241,.55);animation: floatLogo 3s ease-in-out infinite, fadeInOut 6s ease-in-out infinite;" alt="logo"></div>`
        : `<div class="text-center mb-3"><div class="logo logo-hero logo-anim" style="width:${logoSize}px;height:${logoSize}px;border-radius:50%;background:radial-gradient(circle at 30% 30%,#fff,#c7d2fe);box-shadow:0 12px 32px rgba(99,102,241,.55);animation: floatLogo 3s ease-in-out infinite, fadeInOut 6s ease-in-out infinite;"></div></div>`);
      const header = `<div class="card-header text-center" style="background:linear-gradient(135deg,#312e81,#4f46e5);color:#fff;font-weight:700;">تسجيل الدخول</div>`;
       const fields = `<div class="row g-3 justify-content-center" style="max-width:720px;margin:0 auto;">
            <div class="col-md-5"><input id="login_name_inline" class="form-control" placeholder="اسم المستخدم" autocomplete="off" style="box-shadow:inset 0 1px 6px rgba(0,0,0,.08);" onkeydown="triggerLoginOnEnter(event)"></div>
            <div class="col-md-5"><input id="login_pwd_inline" type="password" class="form-control" placeholder="كلمة المرور" autocomplete="new-password" style="box-shadow:inset 0 1px 6px rgba(0,0,0,.08);" onkeydown="if(event.key==='Enter'){performLoginInline();}"></div>
            <div class="col-md-2 d-grid"><button class="btn btn-primary" style="box-shadow:0 10px 24px rgba(79,70,229,.45);" onclick="performLoginInline()">دخول</button></div>
            <div class="col-12"><div id="loginInlineMsg" class="text-danger small mt-2 text-center"></div></div>
            <div class="col-12 text-center"><button class="btn btn-link text-decoration-underline" type="button" onclick="showAdminRecoveryHelper()">نسيت كلمة المرور؟</button></div>
          </div>`;
      const bodyTop = layoutMode==='inline' ? `<div class="d-flex flex-column align-items-center mb-3" style="gap:12px;">
            ${logo}
            <div class="text-muted">أدخل بيانات الدخول للمتابعة</div>
          </div>` : `<div class="text-center text-muted mb-3">أدخل بيانات الدخول للمتابعة</div>`;
      return `<div class="card mb-3" style="max-width:960px;margin:0 auto;box-shadow:0 16px 48px rgba(15,23,42,.28);border-radius:18px;transform:translateZ(0);">
        ${header}
        <div class="card-body" style="background:linear-gradient(145deg,#eef2ff,#e0e7ff);">
          ${bodyTop}
          ${fields}
        </div>
      </div>`;
    }
    function performLoginInline(){
      const name = document.getElementById('login_name_inline').value.trim();
      const pwd = document.getElementById('login_pwd_inline').value;
      const u = (SETTINGS.users || []).find(x=>x.name===name && x.pwd===pwd);
      if(!u){ document.getElementById('loginInlineMsg').textContent = 'بيانات غير صحيحة'; return; }
       CURRENT_USER = {name:u.name, fullName: u.fullName || u.name || '', role:u.role, officerId:u.officerId||null, tabPrivileges: Array.isArray(u.tabPrivileges) ? u.tabPrivileges : [], mustChangePassword: !!u.mustChangePassword};
      u.lastLoginAt = new Date().toISOString();
      sessionStorage.setItem('currentUser', JSON.stringify(CURRENT_USER));
      ACTIVE_SESSIONS = ACTIVE_SESSIONS.filter(s=>!(s.name===u.name && s.role===u.role));
      ACTIVE_SESSIONS.push(CURRENT_USER);
      logActivity('تسجيل دخول', {name: u.name, role: u.role});
      saveAll();
      updateLoginLabel();
      resetIdleTimer();
      if(CURRENT_USER.mustChangePassword){
        showToast('يجب تغيير كلمة المرور فوراً قبل المتابعة','warning');
        renderTab('account');
      } else {
        showToast('مرحباً '+CURRENT_USER.name,'success');
        renderTab('roster');
      }
    }
    function showAdminRecoveryHelper(){
      const container = document.getElementById('loginInlineMsg');
      if(!container) return showToast('تعذر فتح نموذج الاستعادة','danger');
      ADMIN_RECOVERY_CONTEXT = null;
      container.innerHTML = `<div class="alert alert-info">أدخل اسم المستخدم الذي نسيت كلمة مروره.</div>
        <div class="d-flex flex-wrap gap-2" style="align-items:center;">
          <input id="admin_recover_username" class="form-control" placeholder="اسم المستخدم" style="min-width:220px;">
          <button class="btn btn-warning" type="button" onclick="startAdminRecovery()">متابعة</button>
        </div>`;
      showToast('أدخل اسم المستخدم للمتابعة','warning');
    }
    function pickAdminRecoveryDetail(adminUser){
      const linkedOfficer = adminUser.officerId ? officers.find(o=>o.id===adminUser.officerId) : null;
      const candidates = [
        {key:'fullName', label:'الاسم الكامل', value: adminUser.fullName},
        {key:'email', label:'البريد الإلكتروني', value: adminUser.email},
        {key:'phone', label:'رقم الهاتف', value: adminUser.phone},
        {key:'note', label:'ملاحظة الملف', value: adminUser.note},
        {key:'linkedOfficer', label:'اسم الضابط المرتبط', value: linkedOfficer?.name}
      ].filter(item => item.value);
      if(!candidates.length) return null;
      return candidates[Math.floor(Math.random() * candidates.length)];
    }
    function startAdminRecovery(){
      const username = (document.getElementById('admin_recover_username')?.value || '').trim();
      if(!username) return showToast('أدخل اسم المستخدم أولاً','danger');
      const adminIndex = (SETTINGS.users || []).findIndex(u=>u.name===username);
      if(adminIndex < 0) return showToast('اسم المستخدم غير موجود','danger');
      const adminUser = SETTINGS.users[adminIndex];
      if(adminUser.role !== 'admin') return showToast('استعادة كلمة المرور متاحة لحساب المدير فقط','warning');
      const detail = pickAdminRecoveryDetail(adminUser);
      ADMIN_RECOVERY_CONTEXT = detail ? {index: adminIndex, expected: detail.value} : {index: adminIndex, expected: ''};
      const prompt = detail ? `أدخل ${detail.label} لحساب المدير.` : 'أدخل أي معلومة تعريفية من ملف المدير.';
      const container = document.getElementById('loginInlineMsg');
      if(container){
        container.innerHTML = `<div class="alert alert-info">${prompt}</div>
          <div class="d-flex flex-wrap gap-2" style="align-items:center;">
            <input id="admin_recover_hint" class="form-control" placeholder="المعلومة المطلوبة" style="min-width:220px;">
            <input id="admin_recover_newpwd" type="password" class="form-control" placeholder="كلمة مرور جديدة" style="min-width:220px;">
            <button class="btn btn-warning" type="button" onclick="processAdminRecovery()">إعادة تعيين</button>
          </div>`;
      }
      showToast('أجب عن المعلومة المطلوبة لإعادة التعيين','warning');
    }
    function processAdminRecovery(){
      const context = ADMIN_RECOVERY_CONTEXT;
      const adminUser = (SETTINGS.users || [])[context?.index];
      if(!adminUser || adminUser.role!=='admin') return showToast('لم يتم العثور على حساب المدير','danger');
      const hintVal = (document.getElementById('admin_recover_hint')?.value || '').trim().toLowerCase();
      const newPwd = document.getElementById('admin_recover_newpwd')?.value || '';
      if(!hintVal) return showToast('أدخل المعلومة المطلوبة','danger');
      if(!newPwd || newPwd.length < 4) return showToast('أدخل كلمة مرور جديدة صالحة','danger');
      const expected = (context?.expected || '').toString().trim().toLowerCase();
      if(expected){
        const match = expected.includes(hintVal) || hintVal.includes(expected);
        if(!match) return showToast('المعلومة لا تطابق بيانات المدير','danger');
      }
      adminUser.pwd = newPwd;
      adminUser.mustChangePassword = false;
      ADMIN_RECOVERY_CONTEXT = null;
      saveAll();
      logActivity('استعادة كلمة مرور مدير', {admin: adminUser.name});
      showToast('تم تحديث كلمة مرور المدير. يمكنك تسجيل الدخول الآن.','success');
      const container = document.getElementById('loginInlineMsg');
      if(container) container.textContent = 'تم ضبط كلمة المرور، سجّل الدخول بكلمة المرور الجديدة.';
    }
    function signOut(silent=false){
      const prevUser = CURRENT_USER;
      if(CURRENT_USER){ ACTIVE_SESSIONS = ACTIVE_SESSIONS.filter(s=>!(s.name===CURRENT_USER.name && s.role===CURRENT_USER.role)); }
      CURRENT_USER = null;
      sessionStorage.removeItem('currentUser');
      clearIdleTimer();
      saveAll();
      updateLoginLabel();
      renderTab('roster');
      if(prevUser) logActivity('تسجيل خروج', {actor: prevUser.name, actorRole: prevUser.role});
      if(!silent) showToast('تم تسجيل الخروج','info');
    }
    function switchSession(idx){
      const s = ACTIVE_SESSIONS[idx];
      if(!s) return;
      CURRENT_USER = s;
      sessionStorage.setItem('currentUser', JSON.stringify(CURRENT_USER));
      updateLoginLabel();
      renderTab('roster');
      showToast('تم التبديل إلى '+s.name,'success');
      resetIdleTimer();
    }

    function currentUserIndex(){
      if(!CURRENT_USER) return -1;
      return (SETTINGS.users||[]).findIndex(u=>u.name===CURRENT_USER.name);
    }

    function currentUserRecord(){
      const idx = currentUserIndex();
      return idx>=0 ? SETTINGS.users[idx] : null;
    }

    ['click','keydown','mousemove','touchstart','scroll'].forEach(ev=>{
      document.addEventListener(ev, ()=> resetIdleTimer());
    });

   /* ========= Toast ========= */
    function showToast(msg,type='info'){
      const c=document.querySelector('.toast-container');
      const el=document.createElement('div');
      const allowed = ['danger','success','warning','info','request'];
      const cls = allowed.includes(type) ? type : 'info';
      el.className=`toast-item ${cls}`;
      el.textContent = msg;
      c.appendChild(el);
      setTimeout(()=>{ el.style.opacity='0'; setTimeout(()=>el.remove(), 400); }, 3600);
    }
    function safeShowToast(msg,type='info'){ try{ showToast(msg,type); } catch(_) { console.log(msg); } }
    function guardSearchInput(event){
      if(event.key === 'Enter'){
        event.preventDefault();
        event.stopPropagation();
      }
    }
     function renderSessionChips(){ 
     if(!ACTIVE_SESSIONS.length) return '';
      return `<div class="d-flex flex-wrap justify-content-center gap-2">${ACTIVE_SESSIONS.map((s,i)=>`<button class="btn btn-outline-primary btn-sm" onclick="switchSession(${i})">${escapeHtml(s.name)} (${s.role})</button>`).join('')}</div>`;
    }

    function getOfficerRandomBias(id){
      if(!OFFICER_RANDOMNESS[id]) OFFICER_RANDOMNESS[id] = Math.random();
      return OFFICER_RANDOMNESS[id];
    }
    function isOfficerFullyBannedFromDuty(officer, duty){
      return (exceptions || []).some(rule => {
        const dutyMatches = !rule.dutyId || rule.dutyId === duty.id;
        if(!dutyMatches) return false;
        const permanentBlock = (rule.type === 'block' || rule.type === 'deny_duty');
        if(!permanentBlock) return false;
        const targetMatch =
          (rule.targetType === 'officer' && rule.targetId === officer.id) ||
          (rule.targetType === 'dept' && rule.targetId === officer.deptId) ||
          (rule.targetType === 'rank' && rule.targetId === officer.rank);
        if(!targetMatch) return false;
        const hasDateScope = (Array.isArray(rule.weekdays) && rule.weekdays.length) || rule.weekday != null || rule.dayOfMonth || (rule.fromDate && rule.toDate);
        return !hasDateScope;
      });
    }
    function allowedDutyCountForOfficer(officer){
      const allowed = duties.filter(d => !isOfficerFullyBannedFromDuty(officer, d)).length;
      return allowed || duties.length || 1;
    }

    /* ========= Roster Tab ========= */
    function renderRosterTab(){
      if (SETTINGS.authEnabled && !CURRENT_USER) return renderLandingView();
      const viewOnly = isViewOnlyUser();
      const savedMonth = sessionStorage.getItem('activeRosterMonth') || new Date().toISOString().slice(0,7);
      const activeMonth = viewOnly ? new Date().toISOString().slice(0,7) : savedMonth;
      if(viewOnly) sessionStorage.setItem('activeRosterMonth', activeMonth);
      sanitizeRosterAgainstExceptions(activeMonth);
      const editControls = canEditRoster() ? `<button class="btn btn-success me-2" onclick="generateAdvancedRoster()">توزيع ذكي</button>
            <button class="btn btn-outline-primary me-2" onclick="saveRosterAndExport()">حفظ وتصدير الجدول (JSON)</button>
            <button class="btn btn-outline-secondary me-2" onclick="resetRosterForMonth()">تفريغ الشهر</button>` : `<div class="small-muted">وضع عرض فقط</div>`;
      const reportPrintControl = `<button class="btn btn-outline-primary" onclick="printSelectedReportFromRoster()">طباعة التقرير المحدد</button>`;
      return `<div class="card"><div class="card-header">منشئ/محرر جدول الخدمات</div><div class="card-body">
        <div class="row g-2 mb-2">
          <div class="col-md-3"><label class="form-label">شهر الجدول</label><input id="rosterMonth" type="month" class="form-control" value="${activeMonth}" ${viewOnly?'disabled':''} onchange="onRosterMonthChange()"></div>
          <div class="col-md-9 text-end align-self-end">
            ${editControls}
            ${reportPrintControl}
          </div>
        </div>
        <div class="small-muted mb-3">القوائم تعرض فقط الضباط المؤهلين لذلك اليوم/النوع وفق القواعد، وحدود الراحة، وحدود الضباط المخصصة.</div>
        ${duties.map(d=>renderDutySection(d)).join('')}
      </div></div>`;
    }
    function renderLandingView(){
      const logoSize = 288;
      const logo = SETTINGS.logoData ? `<img src="${SETTINGS.logoData}" class="logo logo-hero logo-anim" style="width:${logoSize}px;height:${logoSize}px;border-radius:50%;box-shadow:0 0 48px rgba(99,102,241,.7);animation: floatLogo 3s ease-in-out infinite, fadeInOut 6s ease-in-out infinite;">` : `<div class="logo logo-hero logo-anim" style="width:${logoSize}px;height:${logoSize}px;border-radius:50%;background:radial-gradient(circle at 30% 30%,#fff,#c7d2fe);box-shadow:0 0 48px rgba(99,102,241,.7);animation: floatLogo 3s ease-in-out infinite, fadeInOut 6s ease-in-out infinite;"></div>`;
      if(SETTINGS.loginLayout === 'stacked'){
        return `<div class="d-flex flex-column align-items-center justify-content-center" style="min-height:70vh;gap:18px;">
          <div class="d-flex flex-column align-items-center" style="gap:12px;">
            ${logo}
            <div class="fw-bold fs-4">مرحباً بك في ${escapeHtml(SETTINGS.appName)}</div>
          </div>
          <div style="width:100%;max-width:760px;">${signInInlineFormMarkup(false, 'stacked')}</div>
          ${renderSessionChips()}
        </div>`;
      }
      return `<div class="d-flex flex-column align-items-center justify-content-center" style="min-height:70vh;gap:22px;">
        ${signInInlineFormMarkup(true, 'inline')}
        ${renderSessionChips()}
      </div>`;
    }
    function onRosterMonthChange(){
      if(isViewOnlyUser()) return;
      const m=document.getElementById('rosterMonth').value;
      if(m) sessionStorage.setItem('activeRosterMonth',m);
      renderTab('roster');
    }
    function resetRosterForMonth(){
      if(!canEditRoster()) return safeShowToast('ليس لديك صلاحية للتعديل','danger');
      const month = document.getElementById('rosterMonth')?.value || sessionStorage.getItem('activeRosterMonth');
      if(!month) return safeShowToast('اختر الشهر','danger');
      if(!confirm(`سيتم حذف جدول ${month}. متابعة؟`)) return;
      roster[month] = {};
      saveAll();
      logActivity('تفريغ جدول شهر', {month});
      safeShowToast('تم التفريغ','success');
      renderTab('roster');
    }

    function toggleDutyCollapse(id){ dutyCollapseState[id] = !dutyCollapseState[id]; renderTab('roster'); }
    function toggleSwapBox(id){ dutySwapState[id] = !dutySwapState[id]; renderTab('roster'); }
    function swapDutyDays(dutyId){
      if(!canEditRoster()) return safeShowToast('ليس لديك صلاحية للتعديل','danger');
      const month = document.getElementById('rosterMonth')?.value || sessionStorage.getItem('activeRosterMonth') || new Date().toISOString().slice(0,7);
      const dayA = +document.getElementById(`swap_${dutyId}_a`)?.value;
      const dayB = +document.getElementById(`swap_${dutyId}_b`)?.value;
      if(!dayA || !dayB || dayA===dayB) return safeShowToast('اختر يومين مختلفين للتبديل','warning');
      const [yy,mm] = month.split('-');
      const maxDay = new Date(+yy, +mm, 0).getDate();
      if(dayA<1 || dayB<1 || dayA>maxDay || dayB>maxDay) return safeShowToast('الأيام خارج نطاق الشهر','warning');
      roster[month] = roster[month] || {};
      roster[month][dutyId] = roster[month][dutyId] || [];
      const ensureRow = (day)=>{
        let r = roster[month][dutyId].find(x=>x.day===day);
        if(!r){ r = {day, officerId:null, notes:''}; roster[month][dutyId].push(r); }
        return r;
      };
      const rowA = ensureRow(dayA);
      const rowB = ensureRow(dayB);
      [rowA.officerId, rowB.officerId] = [rowB.officerId, rowA.officerId];
      [rowA.notes, rowB.notes] = [rowB.notes, rowA.notes];
      saveAll();
      safeShowToast('تم التبديل السريع بين اليومين','success');
      renderTab('roster');
    }
    function swapDutyAssignments(dutyId, dayA, dayB){
      if(!canEditRoster()) return safeShowToast('ليس لديك صلاحية للتعديل','danger');
      const month = document.getElementById('rosterMonth')?.value || sessionStorage.getItem('activeRosterMonth') || new Date().toISOString().slice(0,7);
      roster[month] = roster[month] || {};
      roster[month][dutyId] = roster[month][dutyId] || [];
      const ensureRow = (day)=>{
        let r = roster[month][dutyId].find(x=>x.day===day);
        if(!r){ r = {day, officerId:null, notes:''}; roster[month][dutyId].push(r); }
        return r;
      };
      const rowA = ensureRow(dayA);
      const rowB = ensureRow(dayB);
      [rowA.officerId, rowB.officerId] = [rowB.officerId, rowA.officerId];
      [rowA.notes, rowB.notes] = [rowB.notes, rowA.notes];
      saveAll();
      safeShowToast('تم تبديل الضباط بين اليومين بالسحب والإفلات','success');
      renderTab('roster');
    }
    function handleDutyDragStart(event, dutyId, day){
      if(!canEditRoster()) return;
      const payload = { dutyId, day };
      event.dataTransfer.setData('text/plain', JSON.stringify(payload));
      event.dataTransfer.effectAllowed = 'move';
    }
    function allowDutyDrop(event){
      if(!canEditRoster()) return;
      event.preventDefault();
      event.currentTarget.classList.add('drag-over');
    }
    function clearDutyDragOver(event){
      event.currentTarget.classList.remove('drag-over');
    }
    function handleDutyDrop(event, dutyId, day){
      if(!canEditRoster()) return;
      event.preventDefault();
      event.currentTarget.classList.remove('drag-over');
      let payload;
      try { payload = JSON.parse(event.dataTransfer.getData('text/plain')); } catch(_) { payload = null; }
      if(!payload || payload.day == null) return;
      if(payload.dutyId !== dutyId) return safeShowToast('لا يمكن التبديل بين نوعين مختلفين','warning');
      if(+payload.day === +day) return;
      swapDutyAssignments(dutyId, +payload.day, +day);
    }

    function renderDutySection(duty){
      try {
        const month = document.getElementById('rosterMonth')?.value || sessionStorage.getItem('activeRosterMonth') || new Date().toISOString().slice(0,7);
        const monthObj = roster[month] || {};
        const dutyRows = monthObj[duty.id] || [];
        const summaryParts = [];
        if(duty.allowedRanks?.length) summaryParts.push(`مسموح: ${duty.allowedRanks.join('، ')}`);
        summaryParts.push('حدود كل ضابط متاحة في تبويب "حدود الضباط"');
        const rlSummary = summaryParts.filter(Boolean).join(' • ');
        const [yy,mm] = month.split('-');
        const days = new Date(Number(yy), Number(mm), 0).getDate();
        const viewerOfficerId = isViewerUser() ? CURRENT_USER?.officerId : null;
        const viewerDays = viewerOfficerId ? dutyRows.filter(r=>r.officerId===viewerOfficerId).map(r=>r.day).filter(Boolean) : [];
        const daysToRender = isViewerUser() ? viewerDays : Array.from({length: days}, (_, idx) => idx + 1);

        const countsByOfficer = {};
        (dutyRows || []).forEach(r => { if (r.officerId) countsByOfficer[r.officerId] = (countsByOfficer[r.officerId] || 0) + 1; });
        const monthlyCounts = buildOfficerDutyCountsForMonth(month);

        const rowsHtml = daysToRender.map(day => {
          const date = new Date(Number(yy), Number(mm) - 1, day);
          const r = dutyRows.find(x => x.day === day) || {officerId: null, notes: ""};
          const off = officers.find(o => o.id === r.officerId) || null;
          const blocked = off ? getBlockedReasons(off, duty, date) : [];
          const cls = blocked.length ? 'blocked' : '';
          const ttl = blocked.length ? `title="${blocked.join(' | ')}"` : '';

          if(isViewerUser()){
            const displayOfficer = off || (viewerOfficerId ? officers.find(o=>o.id===viewerOfficerId) : null);
            return `<tr class="${cls}" ${ttl}><td class="text-center">${day}</td><td class="text-center">${weekdays[date.getDay()]}</td><td class="text-center">${date.toLocaleDateString('ar-EG')}</td>
              <td>${escapeHtml(displayOfficer?.rank || '')}</td>
              <td>${escapeHtml(displayOfficer?.name || '')}</td>
              <td>${escapeHtml(r.notes || '')}</td></tr>`;
          }

          let eligible = [];
          try {
            eligible = officers.filter(o => isOfficerEligibleForDropdown(o, duty, date, month, r.officerId));
          } catch (err) {
            console.error('Eligibility filter error', err);
		eligible = [];
          }

          if (!eligible.length) eligible = officers.filter(o => !(duty.allowedRanks && duty.allowedRanks.length) || duty.allowedRanks.includes(o.rank));
          if (!eligible.length) eligible = officers.slice();

          eligible.sort((a,b) => {
            const fairnessOrder = compareOfficerFairness(a, b, monthlyCounts);
            if(fairnessOrder !== 0) return fairnessOrder;
            return (countsByOfficer[a.id]||0) - (countsByOfficer[b.id]||0);
          });

          const officerOptions = `<option value="">— اختر —</option>` + eligible.map(o => {
            const deptName = o.deptId ? escapeHtml((departments.find(dd=>dd.id===o.deptId)||{name:''}).name) : '';
            const jtName = o.jobTitleId ? escapeHtml((jobTitles.find(j=>j.id===o.jobTitleId)||{name:''}).name) : '';
            const groupBadge = getDeptGroup(o.deptId)==='external' ? ' • خارجي' : '';
            return `<option value="${o.id}" ${o.id === r.officerId ? 'selected' : ''}>${escapeHtml(o.rank)} ${escapeHtml(o.name)}${deptName? ' • '+deptName:''}${jtName? ' • '+jtName:''}${groupBadge}</option>`;
          }).join('');
          const disabledAttr = canEditRoster() ? '' : 'disabled';
          const dragAttrs = canEditRoster()
            ? `draggable="true" ondragstart="handleDutyDragStart(event, ${duty.id}, ${day})" ondragover="allowDutyDrop(event)" ondragleave="clearDutyDragOver(event)" ondrop="handleDutyDrop(event, ${duty.id}, ${day})"`
            : '';

          return `<tr class="${cls}" ${ttl} ${dragAttrs}><td class="text-center">${day}</td><td class="text-center">${weekdays[date.getDay()]}</td><td class="text-center">${date.toLocaleDateString('ar-EG')}</td>
            <td><select class="form-select form-select-sm" ${disabledAttr} onchange="assignOfficerWithValidation('${month}', ${duty.id}, ${day}, this.value)">${officerOptions}</select></td>
            <td><input class="form-control form-control-sm" ${disabledAttr} value="${escapeHtml(r.notes)}" onchange="updateNote('${month}', ${duty.id}, ${day}, this.value)"></td></tr>`;
        }).join('');

        const collapsed = !!dutyCollapseState[duty.id];
        const showSwap = canEditRoster() && !!dutySwapState[duty.id];
        const swapControls = showSwap ? `<div class="d-flex align-items-center gap-2 mb-2">
              <div>تبديل سريع:</div>
              <input type="number" min="1" max="${days}" class="form-control form-control-sm" style="width:120px" placeholder="اليوم الأول" id="swap_${duty.id}_a">
              <input type="number" min="1" max="${days}" class="form-control form-control-sm" style="width:120px" placeholder="اليوم الثاني" id="swap_${duty.id}_b">
              <button class="btn btn-sm btn-success" onclick="swapDutyDays(${duty.id})">تبديل</button>
            </div>` : '';

        const tableHead = isViewerUser()
          ? '<tr><th>اليوم</th><th>الأسبوع</th><th>التاريخ</th><th>الرتبة</th><th>الضابط</th><th>ملاحظات</th></tr>'
          : '<tr><th>اليوم</th><th>الأسبوع</th><th>التاريخ</th><th>الضابط</th><th>ملاحظات</th></tr>';

        return `<div class="card mb-3"><div class="card-body">
          <div class="d-flex justify-content-between align-items-center mb-2">
            <div><strong>${escapeHtml(duty.name)}</strong> — ${duty.duration}س — راحة ${duty.restHours}س</div>
            <div class="d-flex align-items-center gap-1">
              <button class="btn btn-sm btn-outline-secondary" onclick="toggleDutyCollapse(${duty.id})">${collapsed?'عرض':'إخفاء'}</button>
              ${canEditRoster() ? `<button class="btn btn-sm btn-outline-info" onclick="toggleSwapBox(${duty.id})">${showSwap?'إخفاء التبديل':'تبديل سريع'}</button>` : ''}
              <button class="btn btn-sm btn-outline-primary" onclick="printDuty(${duty.id})">طباعة</button>
              <button class="btn btn-sm btn-outline-secondary" onclick="exportDutyPDF(${duty.id})">حفظ PDF</button>
            </div>
          </div>
          ${rlSummary ? `<div class="mb-2 small-muted">${escapeHtml(rlSummary)}</div>` : ''}
          ${swapControls}
          ${collapsed ? '<div class="text-muted">القسم مطوي</div>' : `${isViewerUser() && !viewerOfficerId ? '<div class="alert alert-warning">لم يتم ربط حسابك بضابط لعرض الأيام المسندة.</div>' : ''}${isViewerUser() && viewerOfficerId && !rowsHtml ? '<div class="alert alert-info">لا توجد أيام مسندة لك في هذا النوع خلال الشهر.</div>' : `<div class="table-responsive"><table class="table table-bordered table-sm"><thead class="table-dark">${tableHead}</thead><tbody>${rowsHtml || `<tr><td colspan=\"${isViewerUser() ? 6 : 5}\" class=\"text-center text-muted\">لا توجد أيام لعرضها</td></tr>`}</tbody></table></div>`}`}
        </div></div>`;
      } catch (err) {
        console.error('renderDutySection error', err);
        return `<div class="alert alert-danger">خطأ عرض المناوبة — تحقق من الكونسول</div>`;
      }
    }

    /* ========= Eligibility Helpers ========= */
    function isOfficerEligibleForDropdown(officer, duty, dateObj, monthStr, currentSelectedId){
      if (!officer) return false;
      if (officer.status && officer.status !== 'active') return false;
      const day = dateObj.getDate();
      const exceptionDecision = evaluateExceptionFor(officer, duty, dateObj);
      // استثناءات الحظر والإزالة تتغلب دائماً على أي تعيين حالي
      if (exceptionDecision.blocked) return false;
      const forced = !!exceptionDecision.forced;
      const lastInfo = getLastAssignmentInfo(monthStr, officer.id, day);

      if (!forced && duty.allowedRanks && duty.allowedRanks.length && !duty.allowedRanks.includes(officer.rank)) {
        if (officer.id !== currentSelectedId) return false;
      }

      const monthRoster = roster[monthStr] || {};
      for (const dId of Object.keys(monthRoster)) {
        const rows = monthRoster[dId] || [];
        if (rows.some(r => r.day === day && r.officerId === officer.id)) {
          if (officer.id !== currentSelectedId) return false;
        }
      }

      const dutyRows = monthRoster[duty.id] || [];
      const assignedCount = dutyRows.filter(r => r.officerId === officer.id).length;
      const totalAssigned = Object.values(monthRoster).reduce((acc, rows)=> acc + (rows||[]).filter(r=>r.officerId===officer.id).length, 0);
      const dutyLimit = getOfficerDutyLimit(officer.id, duty.id);
      const totalLimit = getOfficerTotalLimit(officer.id);

      if (!forced && dutyLimit != null && assignedCount >= dutyLimit && officer.id !== currentSelectedId) return false;
      if (!forced && totalLimit != null && totalAssigned >= totalLimit && officer.id !== currentSelectedId) return false;

      if (!forced && lastInfo && lastInfo.restDays>0) {
        const gap = day - lastInfo.day;
        if (gap <= lastInfo.restDays) {
          if (officer.id !== currentSelectedId) return false;
        }
      }
      if (hasDeptConsecutiveConflict(officer, duty.id, day, monthRoster) && officer.id !== currentSelectedId) return false;
      return true;
    }

    function getBlockedReasons(officer, duty, dateObj){
      const reasons = [];
      const wd = dateObj.getDay();
      for (const rule of (exceptions || [])) {
        try {
          if (Array.isArray(rule.excludedOfficers) && rule.excludedOfficers.includes(officer.id)) continue;
          const dutyMatches = !rule.dutyId || rule.dutyId === duty.id;
          if ((rule.type === 'block' || rule.type === 'deny_duty') && dutyMatches) {
            if (rule.targetType === 'officer' && rule.targetId === officer.id) reasons.push('محظور');
            if (rule.targetType === 'dept' && rule.targetId === officer.deptId) reasons.push('محظور - قسم');
            if (rule.targetType === 'rank' && rule.targetId === officer.rank) reasons.push('محظور - رتبة');
          }
          if (rule.type === 'remove_all' && ((rule.targetType === 'officer' && rule.targetId === officer.id) || (rule.targetType === 'dept' && rule.targetId === officer.deptId) || (rule.targetType === 'rank' && rule.targetId === officer.rank))) {
            reasons.push('مستبعد من كل الجداول');
          }
          if (rule.type === 'vacation' && rule.fromDate && rule.toDate && dutyMatches && ((rule.targetType === 'officer' && rule.targetId === officer.id) || (rule.targetType === 'dept' && rule.targetId === officer.deptId) || (rule.targetType === 'rank' && rule.targetId === officer.rank))) {
            const fr = new Date(rule.fromDate); const to = new Date(rule.toDate); to.setHours(23,59,59);
            if (dateObj >= fr && dateObj <= to) reasons.push('إجازة');
          }
         if (rule.type === 'weekday' && dutyMatches && ruleMatchesWeekday(rule, wd, dateObj)) {
            const label = (Array.isArray(rule.weekdays) && rule.weekdays.length ? rule.weekdays : [rule.weekday]).map(d=>weekdays[+d]).join('، ');
            const modeSuffix = rule.weekdayMode === 'alternate' && Array.isArray(rule.weekdays) && rule.weekdays.length > 1 ? ' (تناوب أسبوعي)' : '';
            if (rule.targetType === 'officer' && rule.targetId === officer.id) reasons.push(`ممنوع يوم ${label}${modeSuffix}`);
            if (rule.targetType === 'dept' && rule.targetId === officer.deptId) reasons.push(`ممنوع يوم ${label}${modeSuffix} للقسم`);
            if (rule.targetType === 'rank' && rule.targetId === officer.rank) reasons.push(`ممنوع يوم ${label}${modeSuffix} للرتبة`);
          }
          if (rule.type === 'force_weekly_only' && dutyMatches) {
            const weekdayLabel = Array.isArray(rule.weekdays) && rule.weekdays.length
              ? rule.weekdays.map(d=>weekdays[+d]).join('، ')
              : weekdays[rule.weekday ?? wd];
            const lbl = `محجوز حصرياً ليوم ${weekdayLabel}${rule.dayOfMonth ? ' ('+rule.dayOfMonth+')' : ''}`;
            if (rule.targetType === 'officer' && rule.targetId === officer.id) reasons.push(lbl);
            if (rule.targetType === 'dept' && rule.targetId === officer.deptId) reasons.push(lbl + ' - القسم');
          }
        } catch (err) { console.error('getBlockedReasons error', err); }
     }
      return reasons;
    }

    function getOfficerIdForDutyDay(rosterSource, dutyId, day){
      const rows = rosterSource?.[dutyId] || [];
      for (let i = 0; i < rows.length; i++) {
        const r = rows[i];
        const rDay = r.day || i + 1;
        if (rDay === day) return r.officerId || null;
      }
      return null;
    }

    function hasDeptConsecutiveConflict(officer, dutyId, day, primaryRoster, secondaryRoster=null){
      if(!officer) return false;
      const deptId = officer.deptId || null;
      const checkDay = (targetDay) => {
        if(targetDay < 1) return false;
        const primary = getOfficerIdForDutyDay(primaryRoster || {}, dutyId, targetDay);
        const fallback = secondaryRoster ? getOfficerIdForDutyDay(secondaryRoster, dutyId, targetDay) : null;
        const resolvedId = primary != null ? primary : fallback;
        if(!resolvedId) return false;
        const other = officers.find(o=>o.id===resolvedId);
        return other && other.deptId === deptId;
      };
      return checkDay(day-1) || checkDay(day+1);
    }

    function isOfficerBannedFromDuty(officer, duty, dateObj){
      if(!officer || !duty) return false;
      const decision = evaluateExceptionFor(officer, duty, dateObj);
      return decision.blocked;
    }

    function assignOfficerWithValidation(monthStr, dutyId, day, officerValue){
      try {
        if(!canEditRoster()) { safeShowToast('ليس لديك صلاحية للتعديل','danger'); renderTab('roster'); return; }
        const officerId = officerValue ? +officerValue : null;
        if (officerId == null) { assignOfficer(monthStr, dutyId, day, null); renderTab('roster'); return; }
        const duty = duties.find(d => d.id === dutyId);
        const officer = officers.find(o => o.id === officerId);
        if (!duty || !officer) { safeShowToast('خطأ في التعيين', 'danger'); return; }
        const dateObj = new Date(+monthStr.split('-')[0], +monthStr.split('-')[1] - 1, day);
        const monthRoster = roster[monthStr] || {};
        if (!isOfficerEligibleForDropdown(officer, duty, dateObj, monthStr, officerId)) {
          safeShowToast('الضابط غير مؤهل لهذا التاريخ/النوع حسب القواعد', 'danger');
          renderTab('roster');
          return;
        }
        if (hasDeptConsecutiveConflict(officer, duty.id, day, monthRoster)) {
          safeShowToast('لا يمكن تعيين ضابطين من نفس القسم ليومين متتاليين في نفس الخدمة', 'danger');
          renderTab('roster');
          return;
        }
        assignOfficer(monthStr, dutyId, day, officerId);
        safeShowToast('تم تعيين الضابط', 'success');
        renderTab('roster');
      } catch (err) {
        console.error('assignOfficerWithValidation error', err);
        safeShowToast('خطأ أثناء التعيين — راجع الكونسول', 'danger');
      }
    }

    function assignOfficer(m, d, y, id){
      if (!roster[m]) roster[m] = {};
      if (!roster[m][d]) roster[m][d] = [];
      const row = roster[m][d].find(r => r.day === y);
      if (row) row.officerId = id || null;
      else roster[m][d].push({day: y, officerId: id || null, notes: ""});
      const dutyObj = duties.find(dt=>dt.id===d);
      const officerObj = officers.find(o=>o.id===id);
      logActivity(id ? 'تعيين خدمة يدوي' : 'إزالة تعيين', {
        month: m,
        dutyName: dutyObj ? dutyObj.name : d,
        day: y,
        target: officerObj ? officerObj.name : 'بدون'
      });
      saveAll();
    }
    function updateNote(m, d, y, n){
      if (!roster[m]) roster[m] = {};
      if (!roster[m][d]) roster[m][d] = [];
      const row = roster[m][d].find(r => r.day === y);
      if (row) row.notes = n;
      else roster[m][d].push({day: y, officerId: null, notes: n});
      const dutyObj = duties.find(dt=>dt.id===d);
      logActivity('تحديث ملاحظة', {month: m, dutyName: dutyObj ? dutyObj.name : d, day: y, note: n});
      saveAll();
    }
    function updateNote(m, d, y, n){
      if (!roster[m]) roster[m] = {};
      if (!roster[m][d]) roster[m][d] = [];
      const row = roster[m][d].find(r => r.day === y);
      if (row) row.notes = n;
      else roster[m][d].push({day: y, officerId: null, notes: n});
      const dutyObj = duties.find(dt=>dt.id===d);
      logActivity('تحديث ملاحظة', {month: m, dutyName: dutyObj ? dutyObj.name : d, day: y, note: n});
      saveAll();
    }

    /* ========= Statistics ========= */
    function prevMonthFrom(monthStr){
      if(!monthStr) return null;
      const [y,m] = monthStr.split('-').map(Number);
      const d = new Date(y, m-1, 1); d.setMonth(d.getMonth() - 1); return d.toISOString().slice(0,7);
    }
    function computeRankStats(){ const counts = {}; for(const r of ranks) counts[r] = 0; for(const o of officers){ if(!o || !o.rank) continue; counts[o.rank] = (counts[o.rank] || 0) + 1; } return counts; }
    function renderRankStatsBar(){ const stats = computeRankStats(); const total = officers.length; return `<div class="mb-3 d-flex flex-wrap gap-2" style="align-items:center"><div class="badge-stat"><div style="font-size:12px">إجمالي الضباط</div><div class="fw-bold">${total}</div></div>${Object.keys(stats).map(rank => `<div class="badge-stat"><div style="font-size:12px">${escapeHtml(rank)}</div><div class="fw-bold">${stats[rank]}</div></div>`).join('')}</div>`; }
    function computeDeptGroupStats(){
      const groupDefs = SETTINGS.deptGroups || [{key:'internal', name:'المجموعة الداخلية'},{key:'external', name:'المجموعة الخارجية'}];
      const stats = {};
      groupDefs.forEach(g=>{
        stats[g.key] = { name: g.name, total: 0, ranks: {} };
        ranks.forEach(r=>{ stats[g.key].ranks[r] = 0; });
      });
      officers.forEach(o=>{
        if(!o) return;
        const groupKey = (departments.find(d=>d.id===o.deptId)?.groupKey) || 'internal';
        if(!stats[groupKey]) stats[groupKey] = { name: groupKey, total: 0, ranks: {} };
        if(!stats[groupKey].ranks[o.rank]) stats[groupKey].ranks[o.rank] = 0;
        stats[groupKey].total += 1;
        if(o.rank) stats[groupKey].ranks[o.rank] += 1;
      });
      return stats;
    }
    function renderDeptGroupStatsTable(){
      const stats = computeDeptGroupStats();
      const rankCols = ranks.map(r=>`<th>${escapeHtml(r)}</th>`).join('');
      const rows = Object.keys(stats).map(key=>{
        const group = stats[key];
        const rankCells = ranks.map(r=>`<td>${group.ranks[r] || 0}</td>`).join('');
        return `<tr><td>${escapeHtml(group.name)}</td><td>${group.total}</td>${rankCells}</tr>`;
      }).join('');
      return `<div class="table-responsive"><table class="table table-sm table-bordered"><thead class="table-dark"><tr><th>مجموعة الأقسام</th><th>الإجمالي</th>${rankCols}</tr></thead><tbody>${rows}</tbody></table></div>`;
    }
   function buildOfficerDutyCountsForMonth(month){
      const result = {};
      const monthRoster = roster[month] || {};
      const [yy, mm] = month.split('-');
      const bannedCarry = {};
      for(const dIdStr of Object.keys(monthRoster)){
        const dId = +dIdStr;
        (monthRoster[dIdStr] || []).forEach(r => {
          if(!r.officerId) return;
          const officer = officers.find(o=>o.id===r.officerId);
          const dutyObj = duties.find(d=>d.id===dId);
          if(!officer || !dutyObj) return;
          if(!result[officer.id]) result[officer.id] = { total: 0, byDuty: {} };
          result[officer.id].total = (result[officer.id].total || 0) + 1;
          const dateObj = new Date(+yy, +mm - 1, r.day || 1);
          const banned = isOfficerBannedFromDuty(officer, dutyObj, dateObj);
          if(banned){
            if(!bannedCarry[officer.id]) bannedCarry[officer.id] = {};
            bannedCarry[officer.id][dId] = (bannedCarry[officer.id][dId] || 0) + 1;
            return;
          }
          result[officer.id].byDuty[dId] = (result[officer.id].byDuty[dId] || 0) + 1;
        });
      }
      Object.keys(bannedCarry).forEach(officerId => {
        const carry = bannedCarry[officerId] || {};
        const existingDuties = Object.keys(result[officerId]?.byDuty || {}).map(Number);
        const fallbackTargets = duties.map(d=>d.id);
        const targets = existingDuties.length ? existingDuties : fallbackTargets;
        Object.entries(carry).forEach(([bannedDutyId, count]) => {
          targets.filter(tid => +tid !== +bannedDutyId).forEach(tid => {
            result[officerId].byDuty[tid] = (result[officerId].byDuty[tid] || 0) + count;
          });
        });
      });
      return result;
    }

    function renderOfficersArchiveTable(){
      const archived = officers.filter(o=> (o.status || 'active') !== 'active');
      if(!archived.length) return '<div class="text-muted">لا يوجد ضباط مؤرشفون حالياً.</div>';
      const rows = archived
        .slice()
        .sort((a,b)=> String(b.archivedAt || '').localeCompare(String(a.archivedAt || '')))
        .map(o=>{
          const deptName = o.deptId ? (departments.find(d=>d.id===o.deptId)?.name || '') : '';
          const archivedAt = o.archivedAt ? new Date(o.archivedAt).toLocaleDateString('ar-EG') : '';
          const [, mid] = splitBadgeParts(o.badge);
          return `<tr>
            <td>${escapeHtml(o.name)}</td>
            <td>${escapeHtml(o.rank)}</td>
            <td>${escapeHtml(o.badge || '')}</td>
            <td>${escapeHtml(mid || '')}</td>
            <td>${escapeHtml(deptName)}</td>
            <td>${renderStatusBadge(o.status || 'active')}</td>
            <td>${escapeHtml(archivedAt)}</td>
          </tr>`;
        }).join('');
      return `<div class="table-responsive"><table class="table table-sm table-bordered align-middle"><thead class="table-dark"><tr><th>الاسم</th><th>الرتبة</th><th>رقم الأقدمية</th><th>الأوسط</th><th>القسم السابق</th><th>الحالة</th><th>تاريخ الأرشفة</th></tr></thead><tbody>${rows}</tbody></table></div>`;
    }
    function renderStatsArchiveTab(){
      if(SETTINGS.authEnabled && !CURRENT_USER) return signInInlineFormMarkup() + '<div class="alert alert-warning">سجل الدخول للوصول للإحصاءات.</div>';
      const savedMonth = document.getElementById('rosterMonth')?.value || sessionStorage.getItem('activeRosterMonth') || new Date().toISOString().slice(0,7);
      const monthsAvailable = Object.keys(roster).sort().reverse();
      const monthOptions = [`<option value="${savedMonth}">Current: ${savedMonth}</option>`].concat(monthsAvailable.map(m=>`<option value="${m}">${m}</option>`)).join('');
      return `<div class="card"><div class="card-header d-flex justify-content-between align-items-center"><div>الإحصاءات والأرشيف</div><div><button class="btn btn-sm btn-outline-primary" onclick="renderTab('roster')">عودة</button></div></div>
        <div class="card-body">
          <h5>إحصاء مجموعات الأقسام</h5>
          <div id="deptGroupStats">${renderDeptGroupStatsTable()}</div>
          <hr/>
          <h5>جدول الواجبات — احصاء شهري</h5>
          <div class="row g-2 mb-2">
            <div class="col-md-3"><label class="form-label">شهر</label><select id="statsMonth" class="form-select">${monthOptions}</select></div>
            <div class="col-md-9 text-end align-self-end">
              <button class="btn btn-success me-2" onclick="renderStatsForSelectedMonth()">عرض الإحصاء</button>
              <button class="btn btn-outline-secondary" onclick="loadAndPreviewPreviousMonth(document.getElementById('statsMonth').value)">معاينة/طباعة الشهر السابق</button>
            </div>
          </div>
          <div id="statsResult"><div class="text-muted">اضغط "عرض الإحصاء" لعرض عدد الواجبات لكل ضابط في الشهر المحدد</div></div>
          <div id="archivePreview" class="mt-3"></div>
          <hr/>
          <h5>أرشيف الضباط (متقاعد/منقول)</h5>
          <div id="officerArchiveTable">${renderOfficersArchiveTable()}</div>
        </div></div>`;
    }
    function renderStatsForSelectedMonth(){
      const month = document.getElementById('statsMonth')?.value || sessionStorage.getItem('activeRosterMonth') || new Date().toISOString().slice(0,7);
      const counts = buildOfficerDutyCountsForMonth(month);
      const header = `<div class="mb-2"><strong>إحصاء لشهر ${month}</strong></div>`;
      const rows = [];
      const dutyNames = duties.map(d=>({id:d.id, name:d.printLabel || d.name}));
      const headCols = `<th>الضابط</th><th>الرتبة</th><th>القسم</th><th>الإجمالي</th>${dutyNames.map(d=>`<th>${escapeHtml(d.name)}</th>`).join('')}`;
      for(const o of officers){
        const c = counts[o.id] || { total:0, byDuty: {} };
        const dutyCells = dutyNames.map(d => `<td>${c.byDuty[d.id] || 0}</td>`).join('');
        const deptName = o.deptId ? escapeHtml((departments.find(dd=>dd.id===o.deptId)||{name:''}).name) : '';
        rows.push(`<tr><td>${escapeHtml(o.name)}</td><td>${escapeHtml(o.rank)}</td><td>${deptName}</td><td>${c.total}</td>${dutyCells}</tr>`);
      }
      const table = `<div class="table-responsive"><table class="table table-sm table-bordered"><thead class="table-dark"><tr>${headCols}</tr></thead><tbody>${rows.join('')}</tbody></table></div>`;
      document.getElementById('statsResult').innerHTML = header + table;
    }
    function loadAndPreviewPreviousMonth(currentMonth){
      const prev = prevMonthFrom(currentMonth);
      const preview = document.getElementById('archivePreview');
      preview.innerHTML = '';
      if(!prev){ preview.innerHTML = `<div class="alert alert-warning">تعذر حساب الشهر السابق</div>`; return; }
      if(!roster[prev]) { preview.innerHTML = `<div class="alert alert-info">لا يوجد جدول محفوظ للشهر السابق (${prev}).</div>`; return; }
      const html = duties.map(d => buildDutyInnerHtml(d.id, prev)).join('<hr/>');
      preview.innerHTML = `<div class="mb-2 d-flex justify-content-between"><div><strong>معاينة جدول ${prev}</strong></div><div><button class="btn btn-primary btn-sm" onclick="printArchive('${prev}')">طباعة هذا الشهر</button></div></div>` + html;
    }
    function printArchive(month){
      if(!roster[month]) { safeShowToast('لا يوجد جدول للطباعة لهذا الشهر', 'danger'); return; }
      const ids = duties.map(d=>d.id);
      printReportPdfStyle(ids, month, `Roster_Archive_${month}.pdf`);
    }

    /* ========= Admin / Backup ========= */
    function getAdminSubTab(){
      return sessionStorage.getItem('adminSubTab') || 'overview';
    }
    function setAdminSubTab(id){
      sessionStorage.setItem('adminSubTab', id);
      renderTab('admin');
    }
    function renderAdminTab(){
      if(SETTINGS.authEnabled && !CURRENT_USER) return signInInlineFormMarkup() + '<div class="alert alert-warning">سجّل الدخول للوصول للوحة التحكم.</div>';
      const summary = `ضباط: ${officers.length} • أقسام: ${departments.length} • مسميات: ${jobTitles.length} • أنواع خدمات: ${duties.length}`;
      const pendingReq = supportRequests.filter(r=>r.status==='pending').length;
      const authSummary = `مستخدمون: ${(SETTINGS.users||[]).length} • جلسات مفعلة: ${ACTIVE_SESSIONS.length} • طلبات مساعدة قيد الانتظار: ${pendingReq}`;
      const active = getAdminSubTab();
      const tabs = [
        {id:'overview', label:'نظرة عامة'},
        {id:'backup', label:'النسخ والتهيئة'},
        {id:'users', label:'المستخدمون'},
        {id:'viewers', label:'مستخدمي المشاهدة'},
        {id:'requests', label:'طلبات الدعم'},
        {id:'activity', label:'سجل النشاط'}
      ];
      const tabButtons = `<div class="sub-tabs">${tabs.map(t=>`<button class="sub-tab-btn ${active===t.id?'active':''}" onclick="setAdminSubTab('${t.id}')">${t.label}</button>`).join('')}</div>`;
      const sections = {
        overview: `
          <div class="mb-3"><strong>ملخص البيانات:</strong> ${summary}</div>
          <div class="small-muted mb-3">${authSummary}</div>
        `,
        backup: `
          <div class="row g-3">
            <div class="col-md-6">
              <div class="border rounded p-3 h-100">
                <h6>نسخ احتياطي واستعادة</h6>
                <p class="small text-muted">قم بحفظ نسخة كاملة من الضباط، الأقسام، الرتب، القواعد، والإعدادات.</p>
                <button class="btn btn-primary mb-2" onclick="exportFullBackup()">تنزيل نسخة احتياطية</button>
                <div class="form-text">استيراد نسخة:</div>
                <input type="file" class="form-control" accept="application/json" onchange="importFullBackup(this.files[0])">
              </div>
            </div>
            <div class="col-md-6">
              <div class="border rounded p-3 h-100">
                <h6>إدارة التخزين والجلسة</h6>
                <p class="small text-muted">استخدم هذه الخيارات بحذر، قد تفقد البيانات المحلية.</p>
                <button class="btn btn-warning mb-2" onclick="clearSessionOnly()">مسح جلسة المستخدم الحالي</button><br/>
                <button class="btn btn-outline-danger" onclick="resetAllData()">تفريغ كل البيانات وإعادتها للوضع الافتراضي</button>
              </div>
            </div>
            <div class="col-md-12">
              <div class="border rounded p-3 h-100 mt-3">
                <h6>تنظيف وتحكم سريع بالجداول</h6>
                <p class="small text-muted">استخدم هذه الأدوات لإدارة الجداول بسرعة دون التأثير على الضباط أو الرتب.</p>
                <div class="mb-2">الشهر النشط: <strong>${sessionStorage.getItem('activeRosterMonth') || new Date().toISOString().slice(0,7)}</strong></div>
                <button class="btn btn-outline-primary mb-2" onclick="adminClearActiveMonth()">مسح جدول الشهر النشط</button><br/>
                <button class="btn btn-outline-secondary mb-2" onclick="adminRestoreDefaults()">استعادة إعدادات الطباعة والشعار الافتراضية</button><br/>
                <button class="btn btn-outline-danger" onclick="adminClearAllRosters()">حذف كل الجداول المخزنة</button>
                <button class="btn btn-outline-primary mt-2" onclick="exportArchive()">تصدير الأرشيف</button>
              </div>
            </div>
          </div>
        `,
        users: `
          <div class="border rounded p-3 h-100">
            <h6>إدارة المستخدمين (صلاحيات وأمن)</h6>
            <div class="row g-2 mb-2">
              <div class="col-md-4"><label class="form-label">مدة الخمول قبل تسجيل الخروج (دقائق)</label><input id="admin_idle_minutes" type="number" class="form-control" value="${SETTINGS.idleLogoutMinutes||0}" min="0" onchange="saveIdleMinutes(this.value)"></div>
              <div class="col-md-8 d-flex align-items-end small-muted">اضبط 0 لتعطيل الخروج التلقائي. لا يتم حفظ الجلسات في التخزين الدائم لزيادة الأمان.</div>
            </div>
            <div class="row g-2 mb-2">
              <div class="col-md-3"><input id="admin_user_name" class="form-control" placeholder="اسم المستخدم"></div>
              <div class="col-md-3"><input id="admin_user_pwd" class="form-control" placeholder="كلمة المرور"></div>
              <div class="col-md-3"><input id="admin_user_fullName" class="form-control" placeholder="الاسم الكامل"></div>
              <div class="col-md-3"><input id="admin_user_email" type="email" class="form-control" placeholder="البريد الإلكتروني"></div>
            </div>
            <div class="row g-2">
              <div class="col-md-2"><select id="admin_user_role" class="form-select" onchange="updatePrivilegeControlState()"><option value="admin">Admin</option><option value="editor">Editor</option><option value="user">User</option><option value="viewer">Viewer</option></select></div>
              <div class="col-md-2"><select id="admin_user_officer" class="form-select"><option value="">ربط بضابط</option>${officers.map(o=>`<option value="${o.id}">${escapeHtml(o.name)} (${escapeHtml(o.rank)})</option>`).join('')}</select></div>
              <div class="col-md-2"><input id="admin_user_phone" class="form-control" placeholder="هاتف"></div>
              <div class="col-md-3"><input id="admin_user_note" class="form-control" placeholder="ملاحظات"></div>
              <div class="col-md-3 text-end"><button class="btn btn-success" onclick="adminAddUser()">إضافة/تعديل</button></div>
            </div>
            <div class="mt-2" id="adminTabPrivileges">
              <label class="form-label">صلاحيات التبويبات للمستخدمين/المحررين</label>
              <div class="compact-checks">
                ${TAB_OPTIONS.map(tab=>`<label class="form-check"><input class="form-check-input admin-privilege-chk" type="checkbox" value="${tab.id}"><span>${escapeHtml(tab.label)}</span></label>`).join('')}
              </div>
              <div class="small text-muted mt-1">عند تحديد الصلاحيات يتم اعتمادها بدلاً من الصلاحيات الافتراضية للدور (User/Editor).</div>
            </div>
            <div class="small text-muted mt-1">الربط بضابط يفعّل الصلاحيات الفردية ويظهر الاسم في الشريط العلوي.</div>
            <div class="mt-3" id="admin_usersTable">${renderUsersTable()}</div>
          </div>
        `,
        viewers: `
          <div class="border rounded p-3 h-100">
            <div class="d-flex justify-content-between align-items-center mb-2">
              <h6 class="mb-0">مستخدمي المشاهدة للضباط</h6>
              <div>
                <button class="btn btn-sm btn-outline-primary me-2" onclick="adminGenerateViewerUsersForAll(false)">إنشاء حسابات المشاهدة</button>
                <button class="btn btn-sm btn-outline-warning me-2" onclick="adminGenerateViewerUsersForAll(true)">تحديث كلمات المرور</button>
                <button class="btn btn-sm btn-outline-secondary" onclick="exportViewerUsersCSV()">تصدير CSV</button>
              </div>
            </div>
            <div class="small text-muted mb-2">ينشئ النظام حساب مشاهدة لكل ضابط لعرض أيامه المسندة فقط.</div>
            <div id="viewerUsersSheet">${renderViewerUsersSheet()}</div>
          </div>
        `,
        requests: `
          <div class="border rounded p-3 h-100">
            <div class="d-flex justify-content-between align-items-center mb-2">
              <h6 class="mb-0">رسائل وطلبات المستخدمين</h6>
              <div class="small text-muted">تظهر طلبات استعادة الدخول والتنبيهات المرسلة للإدارة.</div>
            </div>
            <div class="d-flex justify-content-between align-items-center mb-2">
              <div class="badge-stat">بانتظار الإجراء: ${supportRequests.filter(r=>r.status==='pending').length}</div>
              <div class="text-end">
                <button class="btn btn-sm btn-outline-secondary me-2" onclick="refreshSupportRequestsTable()">تحديث</button>
                <button class="btn btn-sm btn-outline-danger" onclick="clearResolvedRequests()">أرشفة الطلبات المكتملة</button>
              </div>
            </div>
            <div id="supportRequestsTable">${renderSupportRequestsTable()}</div>
          </div>
        `,
        activity: `
          <div class="border rounded p-3 h-100">
            <div class="d-flex justify-content-between align-items-center mb-2">
              <h6 class="mb-0">سجل النشاط</h6>
              <div>
                <button class="btn btn-sm btn-outline-secondary me-2" onclick="exportActivityLog()">تصدير السجل</button>
                <button class="btn btn-sm btn-outline-danger" onclick="clearActivityLog()">مسح السجل</button>
              </div>
            </div>
            <div class="row g-2 mb-2">
              <div class="col-md-3"><input id="act_filter_action" class="form-control" placeholder="بحث بالحدث" oninput="refreshActivityLogTable()" onkeydown="guardSearchInput(event)"></div>
              <div class="col-md-3"><select id="act_filter_user" class="form-select" onchange="refreshActivityLogTable()"><option value="">كل المستخدمين</option>${Array.from(new Set(activityLog.map(l=>l.user||'مجهول'))).map(u=>`<option value="${escapeHtml(u)}">${escapeHtml(u)}</option>`).join('')}</select></div>
              <div class="col-md-3"><input id="act_filter_from" type="date" class="form-control" onchange="refreshActivityLogTable()"></div>
              <div class="col-md-3"><input id="act_filter_to" type="date" class="form-control" onchange="refreshActivityLogTable()"></div>
            </div>
            <div class="small text-muted mb-2">عرض لما تم، متى تم، ومن قام به (يمكن التصفية حسب الحدث، المستخدم، أو الفترة الزمنية).</div>
            <div id="activityLogTableArea">${renderActivityLogTable()}</div>
          </div>
        `
      };
      return `<div class="card">
        <div class="card-header d-flex justify-content-between align-items-center">
          <div>لوحة التحكم والنسخ الاحتياطية</div>
          <div class="text-muted small">${new Date().toLocaleString('ar-EG')}</div>
       </div>
        <div class="card-body">
          ${tabButtons}
          ${sections[active] || ''}
        </div>
      </div>`;
    }
    function renderActivityLogTable(){
      if(!activityLog.length) return '<div class="text-muted">لا يوجد نشاط مسجل بعد</div>';
      const userFilter = document.getElementById('act_filter_user')?.value || '';
      const actionFilter = (document.getElementById('act_filter_action')?.value || '').toLowerCase();
      const fromVal = document.getElementById('act_filter_from')?.value;
      const toVal = document.getElementById('act_filter_to')?.value;
      const fromDate = fromVal ? new Date(fromVal) : null;
      const toDate = toVal ? new Date(toVal) : null;
      const rows = activityLog.slice(-500).filter(entry=>{
        if(userFilter && entry.user !== userFilter) return false;
        if(actionFilter && !(entry.action||'').toLowerCase().includes(actionFilter)) return false;
        if(fromDate && entry.ts && new Date(entry.ts) < fromDate) return false;
        if(toDate && entry.ts && new Date(entry.ts) > new Date(toDate+'T23:59:59')) return false;
        return true;
      }).reverse().map((entry, idx)=> {
        const ts = entry.ts ? new Date(entry.ts).toLocaleString('ar-EG') : '';
        return `<tr>
          <td>${idx+1}</td>
          <td>${escapeHtml(entry.action || '')}</td>
          <td>${escapeHtml(entry.user || 'مجهول')}</td>
          <td>${escapeHtml(ts)}</td>
          <td>${escapeHtml(formatActivityDetails(entry))}</td>
        </tr>`;
      }).join('');
      return `<div class="table-responsive"><table class="table table-sm table-bordered align-middle">
        <thead class="table-dark"><tr><th style="width:60px">#</th><th>النشاط</th><th>المستخدم</th><th>التاريخ/الوقت</th><th>تفاصيل</th></tr></thead>
        <tbody>${rows}</tbody>
      </table></div>`;
    }
    function clearActivityLog(){
      if(!confirm('سيتم مسح كل سجلات النشاط. متابعة؟')) return;
      activityLog = [];
      saveAll();
      renderTab('admin');
      showToast('تم مسح سجل النشاط','info');
    }
    function refreshActivityLogTable(){
      const holder = document.getElementById('activityLogTableArea');
      if(holder) holder.innerHTML = renderActivityLogTable();
    }
    function exportActivityLog(){
      const payload = { log: activityLog, exportedAt: new Date().toISOString() };
      downloadJSONFile('activity_log.json', JSON.stringify(payload, null, 2));
      safeShowToast('تم تنزيل سجل النشاط','success');
    }
    function exportFullBackup(){
      const payload = { officers, departments, jobTitles, duties, roster, exceptions, ranks, settings: SETTINGS, meta:{ exportedAt: new Date().toISOString() } };
      downloadJSONFile(`roster_backup_${new Date().toISOString().slice(0,10)}.json`, JSON.stringify(payload, null, 2));
      logActivity('تصدير نسخة احتياطية', {month: sessionStorage.getItem('activeRosterMonth') || new Date().toISOString().slice(0,7)});
      showToast('تم إنشاء النسخة الاحتياطية','success');
    }
     function importFullBackup(file, options = {}){
      if(!file) return;
      const onComplete = typeof options.onComplete === 'function' ? options.onComplete : null;
      const targetTab = options.targetTab || 'admin';
      const reader = new FileReader();
      reader.onload = e => {
        try {
          const data = JSON.parse(e.target.result || '{}');
          officers = Array.isArray(data.officers) ? data.officers : officers;
          departments = Array.isArray(data.departments) ? data.departments : departments;
          jobTitles = Array.isArray(data.jobTitles) ? data.jobTitles : jobTitles;
          duties = Array.isArray(data.duties) ? data.duties : duties;
          roster = data.roster || {};
          exceptions = Array.isArray(data.exceptions) ? data.exceptions : [];
          ranks = Array.isArray(data.ranks) && data.ranks.length ? data.ranks : ranks;
          SETTINGS = Object.assign({}, defaultSettings, data.settings || {});
          CURRENT_USER = null;
          saveAll();
          loadAll();
          if(onComplete){
            onComplete();
          } else {
            renderTab(targetTab);
          }
          logActivity('استيراد نسخة احتياطية', {file: file?.name || 'upload'});
          showToast('تم الاستيراد بنجاح','success');
        } catch(err){
          console.error(err);
          showToast('تعذر قراءة الملف','danger');
        }
      };
      reader.readAsText(file);
    }
    function adminAddRankQuick(){
      const input = document.getElementById('admin_rank_new');
      const name = (input?.value || '').trim();
      if(!name) return showToast('أدخل اسم رتبة', 'danger');
      if(ranks.includes(name)) return showToast('الرتبة موجودة بالفعل', 'danger');
      ranks.push(name);
      saveAll();
      renderTab('admin');
      logActivity('إضافة رتبة', {target: name});
      showToast('تمت إضافة الرتبة', 'success');
    }
    function adminRenameRank(idx, val){
      const name = (val || '').trim();
      if(!name) return showToast('الاسم مطلوب', 'danger');
      if(ranks.some((r,i)=>r===name && i!==idx)) return showToast('اسم مكرر', 'danger');
      const old = ranks[idx];
      ranks[idx] = name;
      officers.forEach(o=>{ if(o.rank===old) o.rank = name; });
      saveAll();
      logActivity('تعديل رتبة', {target: `${old} -> ${name}`});
      showToast('تم التعديل', 'success');
    }
    function adminDeleteRank(idx){
      const name = ranks[idx];
      if(officers.some(o=>o.rank===name)) return showToast('لا يمكن حذف رتبة مستخدمة', 'danger');
      ranks.splice(idx,1);
      saveAll();
      renderTab('admin');
      logActivity('حذف رتبة', {target: name});
      showToast('تم الحذف', 'success');
    }
    function generateViewerUsername(officer){
      const base = `viewer_${officer.badge || officer.id}`;
      let username = base;
      let suffix = 1;
      while((SETTINGS.users||[]).some(u=>u.name===username)){
        username = `${base}_${suffix++}`;
      }
      return username;
    }
    function adminGenerateViewerUsersForAll(resetPasswords=false){
      if(!officers.length) return showToast('لا يوجد ضباط لإنشاء المستخدمين', 'warning');
      showToast(resetPasswords ? 'يتم تحديث كلمات مرور المشاهدين...' : 'يتم إنشاء حسابات المشاهدة...', 'info');
      SETTINGS.users = SETTINGS.users || [];
      let created = 0;
      let updated = 0;
      officers.forEach(officer=>{
        const existingIdx = SETTINGS.users.findIndex(u=>u.role==='viewer' && u.officerId===officer.id);
        const hasNonViewerAccount = SETTINGS.users.some(u => u.officerId === officer.id && u.role !== 'viewer');
        if(existingIdx >= 0){
          if(resetPasswords){
            SETTINGS.users[existingIdx].pwd = generateTempPassword();
            SETTINGS.users[existingIdx].mustChangePassword = false;
            updated++;
          }
          return;
        }
        if(hasNonViewerAccount) return;
        const username = generateViewerUsername(officer);
        const password = generateTempPassword();
        SETTINGS.users.push({
          name: username,
          pwd: password,
          role: 'viewer',
          officerId: officer.id,
          fullName: officer.name || '',
          createdAt: new Date().toISOString(),
          mustChangePassword: false
        });
        created++;
      });
      saveAll();
      renderTab('admin');
      logActivity('إنشاء مستخدمي مشاهدة', {created, updated});
      showToast(`تم إنشاء ${created} مستخدم، وتحديث ${updated} كلمات مرور`, 'success');
    }
    function exportViewerUsersCSV(){
      const viewers = getViewerUsers();
      if(!viewers.length) return showToast('لا توجد بيانات لتصديرها', 'warning');
      const header = ['الرتبة','الاسم الكامل','اسم المستخدم','كلمة المرور'].map(csvEscape).join(',');
      const rows = viewers.map(v=>[v.rank, v.fullName, v.username, v.password].map(csvEscape).join(',')).join('\n');
      downloadCSVFile(`viewer_users_${new Date().toISOString().slice(0,10)}.csv`, `${header}\n${rows}`);
      showToast('تم تصدير جدول مستخدمي المشاهد', 'success');
    }
    function adminAddUser(){
      const name=document.getElementById('admin_user_name').value.trim();
      const pwd=document.getElementById('admin_user_pwd').value;
      const fullName=document.getElementById('admin_user_fullName').value.trim();
      const email=document.getElementById('admin_user_email').value.trim();
      const role=document.getElementById('admin_user_role').value;
      const officerId=document.getElementById('admin_user_officer').value? +document.getElementById('admin_user_officer').value : null;
      const phone=document.getElementById('admin_user_phone').value;
      const note=document.getElementById('admin_user_note').value;
      if(!name||!pwd) return showToast('ادخل اسم وكلمة مرور','danger');
      SETTINGS.users=SETTINGS.users||[];
      const editIdx = document.getElementById('admin_user_name').dataset.editIndex;
      const existing = (SETTINGS.users||[])[editIdx] || {};
      const selectedTabs = (role === 'user' || role === 'editor') ? getSelectedPrivilegeTabs() : [];
      const payload = {name,pwd,role,officerId,phone,note,fullName,email,tabPrivileges:selectedTabs,createdAt: existing.createdAt || new Date().toISOString()};
      if(editIdx!==undefined && editIdx!==''){
        SETTINGS.users[+editIdx] = payload;
        delete document.getElementById('admin_user_name').dataset.editIndex;
      } else {
        SETTINGS.users.push(Object.assign({mustChangePassword:true}, payload));
      }
      document.getElementById('admin_user_name').value='';
      document.getElementById('admin_user_pwd').value='';
      if(document.getElementById('admin_user_officer')) document.getElementById('admin_user_officer').value='';
      document.getElementById('admin_user_phone').value='';
      document.getElementById('admin_user_note').value='';
      document.getElementById('admin_user_fullName').value='';
      document.getElementById('admin_user_email').value='';
      document.querySelectorAll('.admin-privilege-chk').forEach(chk=>{ chk.checked = false; });
      updatePrivilegeControlState();
      saveAll();
      document.getElementById('admin_usersTable').innerHTML = renderUsersTable();
      const viewerSheet = document.getElementById('viewerUsersSheet');
      if(viewerSheet) viewerSheet.innerHTML = renderViewerUsersSheet();
      logActivity(editIdx!==undefined && editIdx!=='' ? 'تعديل مستخدم' : 'إضافة مستخدم', {target: name, role});
      showToast('تم إضافة المستخدم','success');
    }
    function generateTempPassword(){ return 'Tmp'+Math.random().toString(36).slice(2,6)+'#'+Math.floor(Math.random()*90+10); }
    function adminResetPasswordForUser(idx, requestId=null){
      const u = (SETTINGS.users||[])[idx]; if(!u) return;
      const temp = prompt('أدخل كلمة مرور مؤقتة (سيُطلب تغييرها عند أول دخول):', generateTempPassword());
      if(!temp) return;
      u.pwd = temp;
      u.mustChangePassword = true;
      saveAll();
      if(CURRENT_USER && CURRENT_USER.name === u.name){ CURRENT_USER.mustChangePassword = true; saveAll(); }
      applyNavPermissions();
      document.getElementById('admin_usersTable').innerHTML = renderUsersTable();
      if(requestId) resolveSupportRequest(requestId, 'done', 'تمت إعادة التعيين');
      refreshSupportRequestsTable();
      logActivity('إعادة تعيين كلمة مرور', {target: u.name});
      showToast('تمت إعادة تعيين كلمة المرور','success');
      const viewerSheet = document.getElementById('viewerUsersSheet');
      if(viewerSheet) viewerSheet.innerHTML = renderViewerUsersSheet();
    }
    function saveIdleMinutes(val){
      const minutes = Math.max(0, +val || 0);
      SETTINGS.idleLogoutMinutes = minutes;
      saveAll();
      resetIdleTimer();
      showToast('تم حفظ مدة الخمول','success');
    }
    function adminClearActiveMonth(){
      const month = sessionStorage.getItem('activeRosterMonth') || new Date().toISOString().slice(0,7);
      if(!roster[month]) return showToast('لا يوجد جدول محفوظ لهذا الشهر', 'warning');
      delete roster[month];
      saveAll();
      renderTab('admin');
      logActivity('مسح جدول شهر من لوحة التحكم', {month});
      showToast('تم مسح جدول الشهر النشط','success');
    }
    function adminClearAllRosters(){
      if(!confirm('سيتم حذف كل الجداول المحفوظة. متابعة؟')) return;
      roster = {};
      saveAll();
      renderTab('admin');
      logActivity('حذف كل الجداول', {});
      showToast('تم حذف كل الجداول','success');
    }
    function exportArchive(){
      if(currentRole()!=='admin') return showToast('صلاحية غير كافية','danger');
      const payload = { archivedRoster };
      downloadJSONFile(`roster_archive_${new Date().toISOString().slice(0,10)}.json`, JSON.stringify(payload, null, 2));
      logActivity('تصدير الأرشيف', {});
      showToast('تم تصدير الأرشيف','success');
    }
    function adminRestoreDefaults(){
      if(!confirm('سيتم استعادة الإعدادات الافتراضية للشعار والطباعة. متابعة؟')) return;
      SETTINGS = Object.assign({}, defaultSettings);
      saveAll();
      loadAll();
      renderTab('admin');
      logActivity('استعادة الإعدادات الافتراضية', {});
      showToast('تمت الاستعادة الافتراضية','success');
    }
    function clearSessionOnly(){
      const prevUser = CURRENT_USER;
      CURRENT_USER = null;
      sessionStorage.removeItem('currentUser');
      updateLoginLabel();
      logActivity('مسح جلسة المستخدم', {actor: prevUser?.name, actorRole: prevUser?.role});
      showToast('تم مسح الجلسة الحالية','success');
    }
    function resetAllData(){
      if(!confirm('سيتم حذف كل البيانات المحلية وإعادة ضبط الإعدادات. متابعة؟')) return;
      localStorage.clear();
      sessionStorage.clear();
      officers = [];
      departments = getDefaultDepartments();
      jobTitles = [{id:1,name:"قائد الوحدة",parentId:null,isUpper:true},{id:2,name:"رئيس قسم الشئون الإدارية",parentId:null,isUpper:false}];
      duties = getDefaultDuties();
      roster = {};
      exceptions = [];
      ranks = defaultRanks.slice();
      SETTINGS = Object.assign({}, defaultSettings);
      CURRENT_USER = null;
      activityLog = [];
      logActivity('إعادة ضبط كاملة', {});
      saveAll();
      loadAll();
      renderTab('roster');
      showToast('تمت إعادة الضبط بالكامل','success');
    }

    /* ========= Smart generator ========= */
    function evaluateExceptionFor(officer, duty, dateObj){
      const weekday = dateObj.getDay();
      const dayOfMonth = dateObj.getDate();

      function matchesTarget(rule){
        if (rule.targetType === 'officer') return rule.targetId === officer.id;
        if (rule.targetType === 'dept') return rule.targetId === officer.deptId;
        if (rule.targetType === 'rank') return rule.targetId === officer.rank;
        return false;
      }

      let blocked = false;
      let forced = false;
      let forcedExclusive = false;
      const blockRules = [];
      const forceRules = [];

      const relevantRules = (exceptions || []).filter(r => matchesTarget(r))
        .filter(r => !(Array.isArray(r.excludedOfficers) && r.excludedOfficers.includes(officer.id)));

      for (const rule of relevantRules) {
        try {
          const dutyMatches = !rule.dutyId || rule.dutyId === duty.id;
          const weekdayMatches = ruleMatchesWeekday(rule, weekday, dateObj);
          const dayMatches = rule.dayOfMonth == null || +rule.dayOfMonth === dayOfMonth;

          const applyBlock = (r) => { blocked = true; blockRules.push(r); };
          const applyForce = (r, exclusive=false) => { forced = true; forcedExclusive = forcedExclusive || exclusive; forceRules.push({rule:r, exclusive}); };

          if (rule.type === 'remove_all') { applyBlock(rule); continue; }
          if ((rule.type === 'block' || rule.type === 'deny_duty') && dutyMatches) { applyBlock(rule); continue; }
          if (rule.type === 'vacation' && rule.fromDate && rule.toDate) {
            const fr = new Date(rule.fromDate); const to = new Date(rule.toDate); to.setHours(23,59,59);
            if (dateObj >= fr && dateObj <= to) { applyBlock(rule); continue; }
          }
          if (rule.type === 'weekday' && dutyMatches && weekdayMatches) { applyBlock(rule); continue; }

          if (rule.type === 'force_duty' && dutyMatches && weekdayMatches && dayMatches) { applyForce(rule); continue; }
          if (rule.type === 'force_weekday' && weekdayMatches && dayMatches && dutyMatches) { applyForce(rule); continue; }
          if (rule.type === 'force_weekly_only') {
            if (weekdayMatches && dayMatches && dutyMatches) {
              applyForce(rule, rule.targetType === 'officer');
            } else if (rule.targetType === 'officer') {
              applyBlock(rule); // حصري: يمنع الأيام الأخرى
            }
            continue;
          }
          if (rule.type === 'force' && dutyMatches && weekdayMatches) { applyForce(rule); continue; }
        } catch(err){ console.error('evaluateExceptionFor error', err, rule); }
      }

      const chosenBlock = blockRules.sort((a,b)=> (a.type||'').localeCompare(b.type||''))[0] || null;
      const chosenForce = forceRules.sort((a,b)=> (a.rule.type||'').localeCompare(b.rule.type||''))[0]?.rule || null;
      return {
        blocked,
        forced: forced && !blocked,
        forcedExclusive: forcedExclusive && !blocked,
        rule: blocked ? chosenBlock : chosenForce
      };
    }

    function sanitizeRosterAgainstExceptions(month){
      let changed = false;
      const rosterMonth = roster[month] || {};
      const [yy, mm] = month.split('-');
      Object.keys(rosterMonth).forEach(dutyId => {
        const duty = duties.find(d=>d.id===+dutyId);
        (rosterMonth[dutyId]||[]).forEach(row => {
          if(!row.officerId || !duty) return;
          const officer = officers.find(o=>o.id===row.officerId);
          if(!officer) return;
          const dateObj = new Date(+yy, +mm - 1, row.day || 1);
          const decision = evaluateExceptionFor(officer, duty, dateObj);
          if(decision.blocked){
            row.notes = row.notes || 'مستبعد بالقواعد';
            row.officerId = null;
            changed = true;
         }
        });
      });
      roster[month] = rosterMonth;
      const forcedChanges = enforceForceRulesOnRoster(month);
      if(changed || forcedChanges) saveAll();
    }

    function getForcedTargetsForDate(duty, dateObj){
      const candidates = [];
      const weekday = dateObj.getDay();
      const day = dateObj.getDate();
      for (const rule of (exceptions || [])) {
        if (!(rule.type === 'force' || rule.type === 'force_duty' || rule.type === 'force_weekday' || rule.type === 'force_weekly_only')) continue;
        const dutyMatches = !rule.dutyId || rule.dutyId === duty.id;
        if(!dutyMatches) continue;
        if (!ruleMatchesWeekday(rule, weekday, dateObj)) continue;
        if (rule.dayOfMonth && +rule.dayOfMonth !== day) continue;
        let targets = [];
        if (rule.targetType === 'officer') {
          const o = officers.find(x=>x.id===rule.targetId); if(o) targets=[o];
        } else if (rule.targetType === 'dept') {
          targets = officers.filter(x=>x.deptId === rule.targetId);
        } else if (rule.targetType === 'rank') {
          targets = officers.filter(x=>x.rank === rule.targetId);
        }
        targets.forEach(t => {
          const decision = evaluateExceptionFor(t, duty, dateObj);
          if(decision.forced) candidates.push({officer:t, priority: rule.targetType==='officer'?3:(rule.targetType==='dept'?2:1)});
        });
      }
      candidates.sort((a,b)=> b.priority - a.priority);
      const map = new Map();
      candidates.forEach(c=>{ if(!map.has(c.officer.id)) map.set(c.officer.id, c); });
      return Array.from(map.values()).map(c=>c.officer);
    }

    function enforceForceRulesOnRoster(month){
      let changed = false;
      const rosterMonth = roster[month] || {};
      const [yy, mm] = month.split('-');
      duties.forEach(duty => {
        const rows = rosterMonth[duty.id] || [];
        rows.forEach((row, idx) => {
          const day = row.day || idx + 1;
          const dateObj = new Date(+yy, +mm - 1, day);
          const forcedTargets = getForcedTargetsForDate(duty, dateObj);
          const target = forcedTargets[0];
          if(!target) return;
          if(row.officerId !== target.id){
            row.officerId = target.id;
            row.notes = row.notes || 'تعيين إلزامي حسب القواعد';
            changed = true;
          }
          Object.keys(rosterMonth).forEach(dId => {
            if(+dId === duty.id) return;
            (rosterMonth[dId]||[]).forEach((other, otherIdx) => {
              const otherDay = other.day || otherIdx + 1;
              if(otherDay === day && other.officerId === target.id){
                other.officerId = null;
                other.notes = other.notes || 'أزيل لتعارض مع استثناء إلزامي';
                changed = true;
              }
            });
          });
        });
      });
      roster[month] = rosterMonth;
      return changed;
    }

    function generateAdvancedRoster(){
      try {
        if(!canEditRoster()) { safeShowToast('ليس لديك صلاحية للتعديل','danger'); return; }
        if (SETTINGS.authEnabled && !CURRENT_USER) { safeShowToast('سجل الدخول','danger'); return; }
        const month = document.getElementById('rosterMonth')?.value;
        if (!month) { safeShowToast('اختر الشهر','danger'); return; }
        if (!officers.length) { safeShowToast('أضف ضباطاً','danger'); return; }

        const [yy,mm] = month.split('-');
        const daysInMonth = new Date(Number(yy), Number(mm), 0).getDate();
        const baseMonthRoster = roster[month] || {};
        const prevMonth = prevMonthFrom(month);
        const prevRankTotals = {};
        if(prevMonth && roster[prevMonth]){
          Object.keys(roster[prevMonth] || {}).forEach(dId=>{
            (roster[prevMonth][dId]||[]).forEach(r=>{
              if(!r.officerId) return;
              const o = officers.find(of=>of.id===r.officerId);
              if(!o) return;
              prevRankTotals[o.rank] = (prevRankTotals[o.rank] || 0) + 1;
            });
          });
        }
        const prevRankAvg = (()=>{ const vals = Object.values(prevRankTotals); return vals.length ? vals.reduce((a,b)=>a+b,0)/vals.length : 0; })();
        const newRoster = {};
        const dayUsage = Array.from({length: daysInMonth}, ()=> new Set());
        duties.forEach(d => newRoster[d.id] = Array.from({length: daysInMonth}, (_, i) => ({day: i+1, officerId: null, notes: ""})));

        const officerDutyCounts = {}; const totalOfficerCounts = {}; const rankTotals = {}; const lastAssignedDay = {};
        const lastRestMap = {};
        officers.forEach(o => { officerDutyCounts[o.id] = {}; totalOfficerCounts[o.id] = 0; lastAssignedDay[o.id] = null; lastRestMap[o.id] = {day:null, restDays:0}; });
        duties.forEach(d => { rankTotals[d.id] = {}; ranks.forEach(r => rankTotals[d.id][r] = 0); });

        const placeExistingAssignment = (duty, entry) => {
          const dayIdx = (entry?.day || 0) - 1;
          if(dayIdx < 0 || dayIdx >= daysInMonth) return;
          const slot = newRoster[duty.id][dayIdx];
          if(entry && entry.notes) slot.notes = entry.notes;
          if(!entry.officerId) return;
          const officer = officers.find(o => o.id === entry.officerId);
          if(!officer) return;
          const dateObj = new Date(Number(yy), Number(mm) - 1, entry.day);
          const decision = evaluateExceptionFor(officer, duty, dateObj);
          if(decision.blocked){ slot.notes = entry.notes || 'مستبعد بالقواعد'; return; }
          if(dayUsage[dayIdx].has(officer.id)){ slot.notes = entry.notes || 'تعارض خدمة'; return; }
          slot.officerId = officer.id;
          if(!slot.notes && decision.forced) slot.notes = 'حسب استثناء إلزامي';
          dayUsage[dayIdx].add(officer.id);
          officerDutyCounts[officer.id][duty.id] = (officerDutyCounts[officer.id][duty.id] || 0) + 1;
          totalOfficerCounts[officer.id] = (totalOfficerCounts[officer.id] || 0) + 1;
          rankTotals[duty.id][officer.rank] = (rankTotals[duty.id][officer.rank] || 0) + 1;
          lastAssignedDay[officer.id] = Math.max(lastAssignedDay[officer.id] || 0, entry.day);
          lastRestMap[officer.id] = { day: entry.day, restDays: Math.ceil((duty.restHours||0)/24) };
        };
        duties.forEach(duty => {
          (baseMonthRoster[duty.id] || []).forEach(r => placeExistingAssignment(duty, r));
        });

        function getForceCandidatesFor(duty, dateObj){
          const frs = exceptions.filter(e => (e.type === 'force' || e.type==='force_duty' || e.type==='force_weekday' || e.type==='force_weekly_only') && (!e.dutyId || e.dutyId === duty.id));
          const dayOfWeek = dateObj.getDay();
          const candidates = [];
          for(const fr of frs){
            if(!ruleMatchesWeekday(fr, dayOfWeek, dateObj)) continue;
            if(fr.dayOfMonth && +fr.dayOfMonth !== dateObj.getDate()) continue;
            if(fr.targetType === 'officer' && fr.targetId){
              const o = officers.find(x=>x.id===fr.targetId); if(o) candidates.push(o);
            } else if(fr.targetType === 'dept' && fr.targetId){
              candidates.push(...officers.filter(x=>x.deptId === fr.targetId));
            } else if(fr.targetType === 'rank' && fr.targetId){
              candidates.push(...officers.filter(x=>x.rank === fr.targetId));
            }
          }
          const map = new Map(); candidates.forEach(c=> map.set(c.id, c));
          return Array.from(map.values());
        }
        const fairnessScore = (officer) => {
          const bias = Math.max(0, (prevRankTotals[officer.rank] || 0) - prevRankAvg);
          const availabilityFactor = allowedDutyCountForOfficer(officer) / (duties.length || 1);
          return ((totalOfficerCounts[officer.id] || 0) * availabilityFactor) + bias;
        };
    function isEligibleForGenerator(officer, duty, dateObj, day, usedThisDay){
      if(officer.status && officer.status!=='active') return false;
      const exceptionDecision = evaluateExceptionFor(officer, duty, dateObj);
      if (exceptionDecision.blocked) return false;
      const forced = !!exceptionDecision.forced;
      if(getDeptGroup(officer.deptId)==='external' && !forced) return false;
      const lastInfo = getLastAssignmentInfo(month, officer.id, day, newRoster);

      if (!forced && duty.allowedRanks && duty.allowedRanks.length && !duty.allowedRanks.includes(officer.rank)) return false;
      const monthRoster = baseMonthRoster;
      for (const dId of Object.keys(monthRoster)) {
        const rows = monthRoster[dId] || [];
        if (rows.some(r => r.day === day && r.officerId === officer.id)) return false;
      }
      for(const dId of Object.keys(newRoster)){
        const rows = newRoster[dId] || [];
        if (rows.some(r => r.day === day && r.officerId === officer.id)) return false;
      }
      if (usedThisDay.has(officer.id)) return false;

      const assignedCount = (officerDutyCounts[officer.id]?.[duty.id] || 0);
      const totalAssigned = totalOfficerCounts[officer.id] || 0;
      const dutyLimit = getOfficerDutyLimit(officer.id, duty.id);
      const totalLimit = getOfficerTotalLimit(officer.id);
      if (!forced && dutyLimit != null && assignedCount >= dutyLimit) return false;
      if (!forced && totalLimit != null && totalAssigned >= totalLimit) return false;

      if (!forced && lastInfo && lastInfo.restDays>0){
        const gap = day - lastInfo.day;
        if (gap <= lastInfo.restDays) return false;
      }
      if(hasDeptConsecutiveConflict(officer, duty.id, day, newRoster, monthRoster)) return false;
      return true;
    }

        for (let day = 1; day <= daysInMonth; day++) {
          const usedThisDay = new Set(dayUsage[day - 1] || []);
          for (const duty of duties) {
            const slot = newRoster[duty.id][day - 1];
            if(slot?.officerId){
              usedThisDay.add(slot.officerId);
              continue;
            }
            const dateObj = new Date(Number(yy), Number(mm) - 1, day);
            const forcedCandidates = getForceCandidatesFor(duty, dateObj).filter(o => {
              try { return isEligibleForGenerator(o, duty, dateObj, day, usedThisDay); } catch(_) { return false; }
            });
            if (forcedCandidates.length) {
              forcedCandidates.sort((a,b) => {
                const pa = evaluateExceptionFor(a, duty, dateObj).rule?.targetType === 'officer' ? 3 :
                           evaluateExceptionFor(a, duty, dateObj).rule?.targetType === 'dept' ? 2 : 1;
                const pb = evaluateExceptionFor(b, duty, dateObj).rule?.targetType === 'officer' ? 3 :
                           evaluateExceptionFor(b, duty, dateObj).rule?.targetType === 'dept' ? 2 : 1;
                if (pa !== pb) return pb - pa; // أعلى أولوية أولاً
                const fairnessOrder = compareOfficerFairness(a, b, totalOfficerCounts);
                if(fairnessOrder !== 0) return fairnessOrder;
                return fairnessScore(a) - fairnessScore(b);
              });
              const chosen = forcedCandidates[0];
              newRoster[duty.id][day - 1].officerId = chosen.id;
              if(!newRoster[duty.id][day - 1].notes) newRoster[duty.id][day - 1].notes = 'حسب استثناء إلزامي';
              usedThisDay.add(chosen.id);
              dayUsage[day - 1].add(chosen.id);
              officerDutyCounts[chosen.id][duty.id] = (officerDutyCounts[chosen.id][duty.id] || 0) + 1;
              totalOfficerCounts[chosen.id] = (totalOfficerCounts[chosen.id] || 0) + 1;
              rankTotals[duty.id][chosen.rank] = (rankTotals[duty.id][chosen.rank] || 0) + 1;
              lastAssignedDay[chosen.id] = day;
              lastRestMap[chosen.id] = { day, restDays: Math.ceil((duty.restHours||0)/24) };
              continue;
            }

            let candidates = officers.filter(o => {
              try { return isEligibleForGenerator(o, duty, dateObj, day, usedThisDay); }
              catch(e){ console.error('elig err', e); return false; }
            });
            if (!candidates.length) {
              candidates = officers.filter(o => !(duty.allowedRanks && duty.allowedRanks.length) || duty.allowedRanks.includes(o.rank));
            }
            if (!candidates.length) candidates = officers.slice();

            const frs = (exceptions || []).filter(e => (e.type === 'force' || e.type==='force_duty' || e.type==='force_weekday') && (!e.dutyId || e.dutyId === duty.id));
            if (frs.length) {
              const prefIds = [];
              for (const fr of frs) {
                if (fr.targetType === 'officer') { if (officers.find(x => x.id === fr.targetId)) prefIds.push(fr.targetId); }
                else if (fr.targetType === 'dept') prefIds.push(...officers.filter(x => x.deptId === fr.targetId).map(x => x.id));
                else if (fr.targetType === 'rank') prefIds.push(...officers.filter(x => x.rank === fr.targetId).map(x => x.id));
              }
               const prefSet = Array.from(new Set(prefIds));
              candidates.sort((a,b) => {
                const ai = prefSet.indexOf(a.id) >= 0 ? 0 : 1;
                const bi = prefSet.indexOf(b.id) >= 0 ? 0 : 1;
                if (ai !== bi) return ai - bi;
                const fairnessOrder = compareOfficerFairness(a, b, totalOfficerCounts);
                if(fairnessOrder !== 0) return fairnessOrder;
                return fairnessScore(a) - fairnessScore(b);
              });
            } else {
              candidates.sort((a,b) => {
                const fairnessOrder = compareOfficerFairness(a, b, totalOfficerCounts);
                if(fairnessOrder !== 0) return fairnessOrder;
                return fairnessScore(a) - fairnessScore(b);
              });
            }

            const chosen = candidates[0] || null;
            if (chosen) {
              newRoster[duty.id][day - 1].officerId = chosen.id;
              newRoster[duty.id][day - 1].notes = newRoster[duty.id][day - 1].notes || "";
              usedThisDay.add(chosen.id);
              dayUsage[day - 1].add(chosen.id);
              officerDutyCounts[chosen.id][duty.id] = (officerDutyCounts[chosen.id][duty.id] || 0) + 1;
              totalOfficerCounts[chosen.id] = (totalOfficerCounts[chosen.id] || 0) + 1;
              rankTotals[duty.id][chosen.rank] = (rankTotals[duty.id][chosen.rank] || 0) + 1;
              lastAssignedDay[chosen.id] = day;
              lastRestMap[chosen.id] = { day, restDays: Math.ceil((duty.restHours||0)/24) };
            } else {
              newRoster[duty.id][day - 1].officerId = null;
              newRoster[duty.id][day - 1].notes = "غير متوفر";
            }
          }
        }

        roster[month] = newRoster;
        enforceForceRulesOnRoster(month);
        saveAll();
        logActivity('توزيع ذكي', {month, duties: duties.length, days: daysInMonth});
        safeShowToast('تم توزيع الجدول الذكي وحفظه', 'success');
        renderTab('roster');
      } catch (err) {
        console.error('generateAdvancedRoster error', err);
        safeShowToast('خطأ أثناء توليد الجدول — راجع الكونسول', 'danger');
      }
    }

   /* ========= Print & Export ========= */
    function buildPdfBlobFromCanvas(canvas){
      const pageWidthPt = 595.28; // A4 width in points
      const pageHeightPt = 841.89; // A4 height in points
      const pxPerPt = canvas.width / pageWidthPt;
      const pageHeightPx = Math.floor(pageHeightPt * pxPerPt);
      const pageCanvases = [];
      let renderedHeight = 0;
      while(renderedHeight < canvas.height){
        const tmpCanvas = document.createElement('canvas');
        tmpCanvas.width = canvas.width;
        tmpCanvas.height = Math.min(pageHeightPx, canvas.height - renderedHeight);
        const ctx = tmpCanvas.getContext('2d');
        ctx.drawImage(canvas, 0, renderedHeight, canvas.width, tmpCanvas.height, 0, 0, canvas.width, tmpCanvas.height);
        pageCanvases.push(tmpCanvas);
        renderedHeight += tmpCanvas.height;
      }

      const images = pageCanvases.map(c => ({
        width: c.width,
        height: c.height,
        data: c.toDataURL('image/jpeg', 0.9)
      }));

      function textToUint8(str){ return new TextEncoder().encode(str); }
      function dataUrlToUint8(dataUrl){
        const base64 = dataUrl.split(',')[1];
        const bin = atob(base64);
        const arr = new Uint8Array(bin.length);
        for(let i=0;i<bin.length;i++) arr[i] = bin.charCodeAt(i);
        return arr;
      }

      const chunks = [];
      const offsets = [0];
      const objIndex = [];
      const write = (uint8)=>{ offsets.push(offsets[offsets.length-1] + uint8.length); chunks.push(uint8); };
      const addObject = (content)=>{
        const id = objIndex.length + 1;
        const startOffset = offsets[offsets.length-1];
        const header = textToUint8(`\n${id} 0 obj\n`);
        write(header);
        write(content);
        write(textToUint8('\nendobj\n'));
        objIndex.push(startOffset);
        return id;
      };

      write(textToUint8('%PDF-1.4\n%âãÏÓ\n'));

      const pageIds = [];

      images.forEach((img, idx)=>{
        const imgBinary = dataUrlToUint8(img.data);
        const imgObj = textToUint8(`<< /Type /XObject /Subtype /Image /Width ${img.width} /Height ${img.height} /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode /Length ${imgBinary.length} >>\nstream\n`);
        const imgId = addObject(new Uint8Array([...imgObj, ...imgBinary, ...textToUint8('\nendstream')]));

        const scale = Math.min(pageWidthPt / img.width, pageHeightPt / img.height);
        const drawW = (img.width * scale).toFixed(2);
        const drawH = (img.height * scale).toFixed(2);
        const yPos = (pageHeightPt - img.height * scale).toFixed(2);
        const contentStream = `q\n${drawW} 0 0 ${drawH} 0 ${yPos} cm\n/Im${idx} Do\nQ\n`;
        const contentId = addObject(textToUint8(`<< /Length ${contentStream.length} >>\nstream\n${contentStream}endstream`));

        const pageObj = `<< /Type /Page /Parent 2 0 R /MediaBox [0 0 ${pageWidthPt.toFixed(2)} ${pageHeightPt.toFixed(2)}] /Resources << /ProcSet [/PDF /ImageC] /XObject << /Im${idx} ${imgId} 0 R >> >> /Contents ${contentId} 0 R >>`;
        const pageId = addObject(textToUint8(pageObj));
        pageIds.push(pageId);
      });

      const pagesKids = pageIds.map(id=>`${id} 0 R`).join(' ');
      addObject(textToUint8(`<< /Type /Pages /Count ${pageIds.length} /Kids [${pagesKids}] >>`));
      const pagesId = objIndex.length;
      addObject(textToUint8(`<< /Type /Catalog /Pages ${pagesId} 0 R >>`));
      const catalogId = objIndex.length;

      const xrefStart = offsets[offsets.length-1];
      let xref = `xref\n0 ${catalogId+1}\n0000000000 65535 f \n`;
      const allOffsets = [0, ...objIndex];
      for(const off of allOffsets.slice(1)){
        xref += (off.toString().padStart(10,'0')) + ' 00000 n \n';
      }
      write(textToUint8(xref));
      const trailer = `trailer\n<< /Size ${catalogId+1} /Root ${catalogId} 0 R >>\nstartxref\n${xrefStart}\n%%EOF`;
      write(textToUint8(trailer));

      const totalLength = offsets[offsets.length-1];
      const out = new Uint8Array(totalLength);
      let cursor = 0;
      chunks.forEach(ch=>{ out.set(ch, cursor); cursor += ch.length; });
      return new Blob([out], {type:'application/pdf'});
    }
    function downloadPdfBlob(blob, filename){
      const link = document.createElement('a');
      link.href = URL.createObjectURL(blob);
      link.download = filename || 'roster.pdf';
      link.click();
      setTimeout(()=>{ URL.revokeObjectURL(link.href); }, 800);
    }
    function canvasToPdf(canvas, filename){
      return new Promise((resolve) => {
        const blob = buildPdfBlobFromCanvas(canvas);
        downloadPdfBlob(blob, filename);
        setTimeout(resolve, 800);
      });
    }

    function buildHeaderHtml(){
      if(!SETTINGS.includeHeaderOnPrint) return '';
      return SETTINGS.headerHtml || '';
    }
    function getReportFontFamily(){
      return SETTINGS.reportFontFamily || 'Arial';
    }
    function getReportPrintFontSize(){
      const requested = Number(SETTINGS.printFontSize || defaultSettings.printFontSize || 9);
      return Math.min(requested, REPORT_MAX_FONT_SIZE);
    }
    function buildReportStyles({fontFamily, fontSize, includeBackground = true}){
      return `
          @page{size:A4 portrait;margin:8mm}
          body{font-family:${fontFamily},"Fanan","Sultan bold","PT Bold Heading","Arial","Noto Sans Arabic","Tahoma",sans-serif;margin:0;padding:6mm;${includeBackground ? 'background:#fff;' : ''}}
          .one-duty-page{position:relative;min-height:277mm;page-break-inside:avoid}
             .report-table{width:100%;border-collapse:collapse;table-layout:auto;font-size:${fontSize}px;border:4px double #0f172a}
           .report-table th,.report-table td{border-top:1px solid #0f172a;border-bottom:1px solid #0f172a;border-left:2px solid #0f172a;border-right:2px solid #0f172a;padding:4px 3px;text-align:center;white-space:nowrap}
          .report-table th{background:#d1d5db;color:#111;font-weight:700;border-top:4px double #0f172a;border-bottom:4px double #0f172a;border-left:2px solid #0f172a;border-right:2px solid #0f172a}
          .report-table th.col-index{border-left:4px double #0f172a;border-right:4px double #0f172a}
          .report-table .col-index{background:#d1d5db;font-weight:700;border-left:4px double #0f172a;border-right:4px double #0f172a}
          .report-table .week-separator td{border-top:4px double #0f172a}
          tr{page-break-inside:avoid}
          h2,h3{margin:0}
          .report-header{position:relative;display:flex;justify-content:center;align-items:center;gap:12px;border-bottom:2px solid #0f172a;padding-bottom:6px;margin-bottom:8px}
          .report-meta{display:flex;justify-content:space-between;align-items:center;gap:12px;font-size:11px;color:#111;margin-bottom:6px}
          .report-title{font-size:16px;font-weight:700;text-align:center;color:#111}
          .report-logo{position:absolute;top:0;right:0;display:flex;align-items:center;justify-content:center}
          .report-logo img{max-height:32px;width:auto;height:auto}
          .watermark{position:absolute;top:50%;left:50%;transform:translate(-50%,-50%) rotate(-35deg);font-size:52px;color:#0f172a;opacity:0.08;white-space:pre-wrap;text-align:center;width:160%;pointer-events:none}
          .report-signatures{display:flex;flex-direction:column;gap:12px;margin-top:10px;font-size:12px;line-height:1.5;width:100%;align-items:center;text-align:center}
          .report-signatures-row{display:flex;justify-content:space-between;gap:18px;width:100%}
          .report-signatures-row.edge-spacing{padding-inline:2.5cm}
          .report-signatures-row.centered{justify-content:center}
           .report-signature{flex:0 0 auto;text-align:center;min-height:95px;display:flex;flex-direction:column;justify-content:flex-end;align-items:center;width:fit-content;max-width:100%}
          .report-signature.edge{align-items:center}
          .report-signature.align-right{text-align:center;align-items:center}
           .report-signature.align-left{text-align:center;align-items:center}
          .report-signature.align-right.edge,.report-signature.align-left.edge{align-items:center}
          .report-signature.align-center{text-align:center;align-items:center}
          .report-signature.emphasis{font-size:1.3em;font-weight:800}
          .report-signature .signature-name-block{display:inline-flex;flex-direction:column;align-items:center;align-self:center;width:max-content}
          .report-signature .signature-line{border-top:2px solid #0f172a;margin:12px 0 6px;width:100%}
          .report-signature .signature-place{text-align:right;width:100%;align-self:flex-end}
          .report-signature .signature-name{text-align:center;width:100%}
          .report-signature .signature-title{text-align:center;width:100%}
          .report-signature.edge .signature-line{margin-inline:0}
      `;
    }
    function buildDutyTableHtml(rows, effectiveMonth){
      const [y,m] = effectiveMonth.split('-');
      const days = new Date(+y,+m,0).getDate();
        let table = `<table class="report-table"><thead><tr><th class="col-index">م</th><th>اليوم</th><th>التاريخ</th><th>الرتبة</th><th>الاسم</th><th>الملاحظات</th></tr></thead><tbody>`;
      for(let day=1; day<=days; day++){
        const r = rows.find(x=>x.day===day) || {officerId:null,notes:""};
        const date = new Date(+y,+m-1,day);
        const o = officers.find(x=>x.id===r.officerId) || {rank:'',name:'—'};
        const weekClass = day > 1 && date.getDay() === 6 ? 'week-separator' : '';
        table += `<tr class="${weekClass}"><td class="col-index">${day}</td><td>${weekdays[date.getDay()]}</td><td>${date.toLocaleDateString('ar-EG')}</td><td>${escapeHtml(o.rank)}</td><td>${escapeHtml(o.name)}</td><td></td></tr>`;
      }
      table += '</tbody></table>';
      return table;
    }
     function buildReportHeader(duty, effectiveMonth){
      const monthTitle = new Date(effectiveMonth+'-01').toLocaleDateString('ar-EG',{month:'long',year:'numeric'});
      const title = `${escapeHtml(duty.printLabel || duty.name)}`;
      const logoHtml = SETTINGS.logoData ? `<div class="report-logo"><img src="${SETTINGS.logoData}" alt="logo"></div>` : '';
      return `
        <div class="report-header">
          <div style="flex:1;text-align:center;">
            <div class="report-title">${title}</div>
          </div>
          ${logoHtml}
        </div>
        <div class="report-meta">
          <div>الشهر: ${monthTitle}</div>
          <div>عدد الأيام: ${new Date(+effectiveMonth.split('-')[0], +effectiveMonth.split('-')[1], 0).getDate()}</div>
        </div>
    `;
    }
    function buildSignatureHtml(fontFamily, fontSize){
      const sigs = SETTINGS.signatures || [];
      const signatureFontSize = fontSize ? Math.round(fontSize * 1.2 * 10) / 10 : null;
      const sigFont = fontFamily ? `font-family:${fontFamily},'Arial','Noto Sans Arabic','Tahoma',sans-serif;font-size:${signatureFontSize || fontSize}px;` : '';
     return `<div class="report-signatures" style="${sigFont}">
        <div class="report-signatures-row edge-spacing">
          <div class="report-signature align-left edge">
            <div class="signature-place">${escapeHtml(sigs[0]?.place||'')}</div>
            <div class="signature-name-block">
              <div class="signature-line"></div>
               <div class="signature-name">${escapeHtml(sigs[0]?.name||'')}</div>
            </div>
            <div class="signature-title">${escapeHtml(sigs[0]?.title||'')}</div>
          </div>
           <div class="report-signature align-right edge">
            <div class="signature-place">${escapeHtml(sigs[1]?.place||'')}</div>
            <div class="signature-name-block">
              <div class="signature-line"></div>
              <div class="signature-name">${escapeHtml(sigs[1]?.name||'')}</div>
            </div>
            <div class="signature-title">${escapeHtml(sigs[1]?.title||'')}</div>
          </div>
        </div>
        <div class="report-signatures-row centered">
           <div class="report-signature align-center emphasis">
            <div class="signature-place">${escapeHtml(sigs[2]?.place||'')}</div>
            <div class="signature-name-block">
              <div class="signature-line"></div>
             <div class="signature-name">${escapeHtml(sigs[2]?.name||'')}</div>
            </div>
            <div class="signature-title">${escapeHtml(sigs[2]?.title||'')}</div>
          </div>
        </div>
      </div>`;
    }
   function stripHtmlToText(html){
      const div = document.createElement('div');
      div.innerHTML = html || '';
      return (div.textContent || div.innerText || '').trim();
    }
    function buildReportWatermarkHtml(){
      if(!SETTINGS.includeFooterOnPrint) return '';
      const linkedOfficer = CURRENT_USER?.officerId ? officers.find(o=>o.id===CURRENT_USER.officerId) : null;
      const userName = linkedOfficer?.name || CURRENT_USER?.fullName || CURRENT_USER?.name || 'مستخدم غير مسجل';
      const stamp = new Date().toLocaleString('ar-EG');
      const text = `${userName}\n${stamp}`;
      return `<div class="watermark">${escapeHtml(text)}</div>`;
    }


   function buildDutyHtmlForExport(dutyId, month){
      const effectiveMonth = month || (document.getElementById('rpt_month')?.value || sessionStorage.getItem('activeRosterMonth') || new Date().toISOString().slice(0,7));
      return buildReportHtmlDocument([dutyId], effectiveMonth);
    }
    function triggerIframePrint(html){
      const iframe = document.createElement('iframe');
      iframe.style.position = 'fixed';
      iframe.style.right = '0';
      iframe.style.bottom = '0';
      iframe.style.width = '0';
      iframe.style.height = '0';
      iframe.style.border = '0';
      iframe.setAttribute('aria-hidden', 'true');
      document.body.appendChild(iframe);
      iframe.onload = () => {
        try {
          setTimeout(()=>{ iframe.contentWindow?.focus(); iframe.contentWindow?.print(); iframe.remove(); }, 350);
        } catch(err){
          console.error('iframe print error', err);
          safeShowToast('تعذر بدء الطباعة','danger');
          iframe.remove();
        }
      };
      iframe.srcdoc = html;
    }
    async function ensureHtml2CanvasReady(){
      if(typeof html2canvas === 'function') return true;
      return new Promise((resolve, reject)=>{
        const s = document.createElement('script');
        s.src = 'https://cdn.jsdelivr.net/npm/html2canvas@1.4.1/dist/html2canvas.min.js';
        s.onload = ()=> resolve(true);
        s.onerror = ()=> { safeShowToast('تعذر تحميل مكتبة PDF','danger'); reject(new Error('html2canvas load failed')); };
        document.head.appendChild(s);
      });
    }
    function buildReportCanvasHtml(ids, month){
      const fontFamily = getReportFontFamily();
      const fontSize = getReportPrintFontSize();
      const styles = buildReportStyles({fontFamily, fontSize, includeBackground: true});
      const bodyContent = ids.map(id=> buildDutyInnerHtml(id, month)).join('<div style="page-break-after:always;"></div>');
      return `<style>${styles}</style>${bodyContent}`;
    }
    async function printPdfFromHtml(html, filename){
      const container = document.createElement('div');
      container.style.width = '210mm';
      container.style.padding = '6mm';
      container.style.direction = 'rtl';
      container.innerHTML = html;
      document.body.appendChild(container);
      try{
        await ensureHtml2CanvasReady();
        const canvas = await html2canvas(container, {scale:2, useCORS:true});
        const blob = buildPdfBlobFromCanvas(canvas);
        const url = URL.createObjectURL(blob);
        const win = window.open(url, '_blank');
        if(!win){
          safeShowToast('تعذر فتح نافذة الطباعة','warning');
          return;
        }
        win.onload = () => {
          try { win.focus(); win.print(); } catch(_) {}
        };
        setTimeout(()=>{ URL.revokeObjectURL(url); }, 2000);
      }catch(err){
        console.error(err);
        safeShowToast('فشل إعداد الطباعة','danger');
      }finally{
        container.remove();
      }
    }
    function printReportPdfStyle(ids, month, filename){
      const html = buildReportCanvasHtml(ids, month);
      printPdfFromHtml(html, filename);
    }
    function printDuty(dutyId){
      const month = document.getElementById('rosterMonth')?.value || sessionStorage.getItem('activeRosterMonth') || new Date().toISOString().slice(0,7);
      const duty = duties.find(d=>d.id===dutyId);
      if(!duty){ safeShowToast('تعذر العثور على نوع الخدمة للطباعة','danger'); return; }
      printReportPdfStyle([dutyId], month, `${duty.name}_${month}.pdf`);
    }
    function exportDutyPDF(dutyId){
      const month = document.getElementById('rosterMonth')?.value || sessionStorage.getItem('activeRosterMonth') || new Date().toISOString().slice(0,7);
      const html = buildDutyInnerHtml(dutyId, month);
      const container = document.createElement('div');
      container.style.width = '210mm'; container.style.padding = '6mm'; container.style.direction = 'rtl'; container.innerHTML = html;
      document.body.appendChild(container);
      html2canvas(container, {scale:2, useCORS:true}).then(canvas => canvasToPdf(canvas, `${duties.find(d=>d.id===dutyId).name}_${month}.pdf`).then(()=>{ safeShowToast('تم تنزيل PDF','success'); container.remove(); })).catch(e=>{ safeShowToast('فشل إنشاء PDF','danger'); console.error(e); container.remove(); });
    }

    function saveRosterAndExport(){
      const month = document.getElementById('rosterMonth')?.value || sessionStorage.getItem('activeRosterMonth') || new Date().toISOString().slice(0,7);
      const payload = { month, roster: roster[month] || {}, meta: { generatedBy: CURRENT_USER ? CURRENT_USER.name : 'system', date: new Date().toISOString() } };
      downloadJSONFile(`roster_${month}.json`, JSON.stringify(payload, null, 2));
      logActivity('تصدير جدول JSON', {month});
      safeShowToast('تم حفظ وتصدير الجدول JSON', 'success');
    }

    function buildDutyInnerHtml(dutyId, month){
  	const duty = duties.find(d=>d.id===dutyId);
  const rows = roster[month]?.[dutyId] || [];
  const header = buildHeaderHtml();
  const reportHeader = buildReportHeader(duty, month);
  const fontSize = getReportPrintFontSize();
  const fontFamily = getReportFontFamily();
  const table = buildDutyTableHtml(rows, month);
  const sigs = SETTINGS.signatures || [];
  const signatureHtml = buildSignatureHtml(fontFamily, fontSize);
  const watermark = buildReportWatermarkHtml();
  return `<div class="one-duty-page report-preview-paper">${watermark}<div class="duty-content" style="font-size:${fontSize}px;font-family:${fontFamily},'Arial','Noto Sans Arabic','Tahoma',sans-serif;">${header ? `<div style="margin-bottom:6px;">${header}</div>` : ''}${reportHeader}${table}<div style="margin-top:8px">${signatureHtml}</div></div></div>`;
}

    function printSelectedDuties(){
      const ids = getSelectedDutyIdsInReport(); const month = document.getElementById('rpt_month')?.value || sessionStorage.getItem('activeRosterMonth');
      if(!ids.length) return showToast('اختر نوعاً واحداً أو أكثر للطباعة','danger');
      printReportPdfStyle(ids, month, `Roster_Selected_${month}.pdf`);
    }
    function printSelectedReportFromRoster(){
      let ids = [];
      try { ids = JSON.parse(sessionStorage.getItem('selectedReportDutyIds')) || []; } catch(_) { ids = []; }
      const month = document.getElementById('rosterMonth')?.value || sessionStorage.getItem('activeRosterMonth');
      if(!ids.length){
        ids = duties.map(d=>d.id);
        if(!ids.length) return showToast('لا توجد أنواع خدمات للطباعة','warning');
        showToast('لم يتم تحديد أنواع في التقرير، سيتم طباعة الكل','info');
      }
      printReportPdfStyle(ids, month, `Roster_${month}.pdf`);
    }
    async function saveSelectedDutiesPDF(){
      const ids = getSelectedDutyIdsInReport(); const month = document.getElementById('rpt_month')?.value || sessionStorage.getItem('activeRosterMonth');
      if(!ids.length) return showToast('اختر نوعاً واحداً أو أكثر للحفظ كـ PDF','danger');
      const container = document.createElement('div');
      container.style.width = '210mm'; container.style.padding = '6mm'; container.style.direction = 'rtl';
      container.innerHTML = ids.map(id=> buildDutyInnerHtml(id, month)).join('<div style="page-break-after:always;"></div>');
      document.body.appendChild(container);
      try{
        const canvas = await html2canvas(container, {scale:2, useCORS:true});
        await canvasToPdf(canvas, `Roster_${month}.pdf`);
        showToast('تم إنشاء PDF وتنزيله','success');
      }catch(e){ console.error(e); showToast('فشل إنشاء PDF','danger'); }
      finally{ container.remove(); }
    }
    async function saveAllDutiesPDF(){
      const month = document.getElementById('rosterMonth')?.value || sessionStorage.getItem('activeRosterMonth');
      const ids = duties.map(d=>d.id);
      const container = document.createElement('div');
      container.style.width = '210mm'; container.style.padding = '6mm'; container.style.direction = 'rtl';
      container.innerHTML = ids.map(id=> buildDutyInnerHtml(id, month)).join('<div style="page-break-after:always;"></div>');
      document.body.appendChild(container);
      try{
        const canvas = await html2canvas(container, {scale:2, useCORS:true});
        await canvasToPdf(canvas, `Roster_All_${month}.pdf`);
        showToast('تم إنشاء PDF وتنزيله','success');
      }catch(e){ console.error(e); showToast('فشل إنشاء PDF','danger'); }
      finally{ container.remove(); }
    }
    function printAllDuties(){
      const month = document.getElementById('rosterMonth')?.value || sessionStorage.getItem('activeRosterMonth');
      const ids = duties.map(d=>d.id);
      printReportPdfStyle(ids, month, `Roster_All_${month}.pdf`);
    }


    /* ========= Officers Tab ========= */
    function getOfficerFilterState(){
      return {
        search: sessionStorage.getItem('officerSearch') || ''
      };
    }
    function updateOfficerFilters(){
      const search = document.getElementById('officerSearch')?.value || '';
      sessionStorage.setItem('officerSearch', search);
      const list = document.getElementById('officerList');
      if(list){ list.innerHTML = renderOfficersList(); }
    }
    function resetOfficerFilters(){
      sessionStorage.removeItem('officerSearch');
      renderTab('officers');
    }
    function renderOfficersTab(){
      if(SETTINGS.authEnabled && !CURRENT_USER) return signInInlineFormMarkup() + '<div class="alert alert-warning">سجل الدخول للوصول للضباط.</div>';
      if(SETTINGS.authEnabled && (currentRole()==='user' || currentRole()==='viewer')) return '<div class="alert alert-warning">وضع عرض فقط. تواصل مع المشرف للصلاحيات.</div>';
      const statsBar = renderRankStatsBar();
      const canManage = currentRole()==='admin' || currentRole()==='editor';
      const filters = getOfficerFilterState();
      return `<div class="card"><div class="card-header d-flex justify-content-between align-items-center"><div>الضباط — إدارة</div>
        <div>${canManage?`<button class="btn btn-success btn-sm" onclick="addOfficerForm()">إضافة ضابط</button>`:''}</div></div>
        <div class="card-body">
          ${statsBar}
          <div class="row mb-2 g-2">
            <div class="col-md-6"><input id="officerSearch" class="form-control" placeholder="ابحث بالاسم، الرتبة، الرقم، القسم، أو الحالة" value="${escapeHtml(filters.search)}" oninput="updateOfficerFilters()" onkeydown="guardSearchInput(event)"></div>
            <div class="col-md-2"><button class="btn btn-outline-secondary" onclick="resetOfficerFilters()">إعادة ضبط</button></div>
            <div class="col-md-4 text-end">
              ${canManage?`<button class="btn btn-outline-secondary" onclick="exportOfficersCSV()">تصدير CSV</button>
              <button class="btn btn-outline-secondary ms-2" onclick="exportOfficersFullCSV()">تصدير CSV كامل</button>
              <button class="btn btn-outline-info ms-2" onclick="downloadOfficersTemplateCSV()">تحميل قالب CSV</button>
              <label class="btn btn-outline-primary ms-2 mb-0">استيراد CSV <input type="file" id="officersCsvInput" accept=".csv" style="display:none" onchange="importOfficersCSV(this)"></label>`:''}
            </div>
            </div>
          </div>
          <div class="row">
            <div class="col-md-8"><div id="officerList">${renderOfficersList()}</div></div>
            <div class="col-md-4"><div id="officerDetail">${renderOfficerDetail()}</div></div>
          </div>
        </div></div>
        <div id="officerEditorArea"></div>`;
   }
     function renderOfficersList(){
      const filters = getOfficerFilterState();
      const q = (filters.search || '').toLowerCase();
      const list = officers.filter(o=>{
        if(!q) return true;
        const deptName = o.deptId ? (departments.find(d=>d.id===o.deptId)?.name || '') : '';
        const statusLabel = renderStatusBadge(o.status||'active').replace(/<[^>]+>/g,'');
        const jobTitle = o.jobTitleId ? (jobTitles.find(j=>j.id===o.jobTitleId)?.name || '') : '';
        return o.name?.toLowerCase().includes(q)
          || o.rank?.toLowerCase().includes(q)
          || (o.badge||'').toLowerCase().includes(q)
          || deptName.toLowerCase().includes(q)
          || statusLabel.toLowerCase().includes(q)
          || jobTitle.toLowerCase().includes(q);
      });
      if(!list.length) return `<p class="text-center text-muted py-3">لا توجد نتائج</p>`;
      const rows = list.map(o=>`<tr onclick="SELECTED_OFFICER_ID=${o.id};renderTab('officers');" style="cursor:pointer;${SELECTED_OFFICER_ID===o.id?'background:#eef2ff;':''}">
        <td>${escapeHtml(o.name)}</td>
        <td>${escapeHtml(o.rank)}</td>
        <td>${escapeHtml((departments.find(d=>d.id===o.deptId)||{name:''}).name)}</td>
        <td>${renderStatusBadge(o.status||'active')}</td>
      </tr>`).join('');
      return `<div class="table-responsive"><table class="table table-bordered table-sm"><thead class="table-dark"><tr><th>الاسم</th><th>الرتبة</th><th>القسم</th><th>الحالة</th></tr></thead><tbody>${rows}</tbody></table></div>`;
    }
    function renderStatusBadge(status){
      const label = {active:'نشط', retired:'متقاعد', transferred:'منقول'}[status] || status;
      const cls = status==='retired'?'status-retired':(status==='transferred'?'status-transferred':'status-active');
      return `<span class="badge-status ${cls}">${label}</span>`;
    }
    function renderOfficerDetail(){
      const o = officers.find(x=>x.id===SELECTED_OFFICER_ID) || officers[0];
      if(!o) return '<div class="text-muted">اختر ضابطاً لعرض التفاصيل</div>';
      const canManage = currentRole()==='admin' || currentRole()==='editor';
      const disabledAttr = canManage ? '' : 'disabled';
      const isUpper = jobTitles.find(j=>j.id===o.jobTitleId)?.isUpper;
      const deptDisabled = isUpper ? 'disabled' : disabledAttr;
      const transferDepts = getTransferDepartments();
      const transferInVisible = o.admissionType === 'transferred' || o.status === 'transferred' || o.transferFromDeptId || o.transferDate;
      const transferOutVisible = o.status === 'transferred' || o.transferToDeptId || o.transferToDate;
      const [badgeFirst, badgeSecond, badgeThird] = splitBadgeParts(o.badge);
      const officerTitle = `${escapeHtml(o.rank)} ${escapeHtml(o.name)}`;
      return `<div class="card sticky-panel"><div class="card-header d-flex justify-content-between align-items-center"><div>تفاصيل الضابط</div><div class="badge-stat">${officerTitle}</div></div><div class="card-body">
        <div class="mb-2"><label class="form-label">الاسم</label><input class="form-control" ${disabledAttr} value="${escapeHtml(o.name)}" onchange="updateOfficer(${o.id},'name',this.value)"></div>
        <div class="row g-2">
          <div class="col-md-6"><label class="form-label">الرتبة</label><select class="form-select" ${disabledAttr} onchange="updateOfficer(${o.id},'rank',this.value)">${ranks.map(r=>`<option ${r===o.rank?'selected':''}>${r}</option>`).join('')}</select></div>
          <div class="col-md-6"><label class="form-label">القسم</label><select class="form-select" ${deptDisabled} onchange="updateOfficer(${o.id},'deptId',+this.value||null)"><option value="">— قسم —</option>${getInternalDepartments().map(d=>`<option value="${d.id}" ${d.id===o.deptId?'selected':''}>${escapeHtml(d.name)}</option>`).join('')}</select>${isUpper?'<div class="small-muted">حامل مسمى إدارة عليا - القسم معطل</div>':''}</div>
        </div>
        <div class="row g-2">
          <div class="col-md-6"><label class="form-label">المسمى الوظيفي</label><select class="form-select" ${disabledAttr} onchange="updateOfficer(${o.id},'jobTitleId',+this.value||null)"><option value="">— مسمى —</option>${jobTitles.map(j=>`<option value="${j.id}" ${j.id===o.jobTitleId?'selected':''}>${escapeHtml(j.name)}</option>`).join('')}</select></div>
          <div class="col-md-6"><label class="form-label">رقم الأقدمية</label>
            <div class="d-flex align-items-center gap-1">
              <input id="badge_${o.id}_1" class="form-control" ${disabledAttr} value="${escapeHtml(badgeFirst)}" placeholder="رقم 1" required onchange="updateOfficerBadgeFromInputs(${o.id})" style="min-width:0;">
              <span class="text-muted">/</span>
              <input id="badge_${o.id}_2" class="form-control" ${disabledAttr} value="${escapeHtml(badgeSecond)}" placeholder="رقم 2" required onchange="updateOfficerBadgeFromInputs(${o.id})" style="min-width:0;">
              <span class="text-muted">/</span>
              <input id="badge_${o.id}_3" class="form-control" ${disabledAttr} value="${escapeHtml(badgeThird)}" placeholder="اختياري" onchange="updateOfficerBadgeFromInputs(${o.id})" style="min-width:0;">
            </div>
          </div>
        </div>
        <div class="row g-2">
          <div class="col-md-6"><label class="form-label">هاتف</label><input class="form-control" ${disabledAttr} value="${escapeHtml(o.phone||'')}" onchange="updateOfficer(${o.id},'phone',this.value)"></div>
          <div class="col-md-6"><label class="form-label">الجنس</label><select class="form-select" ${disabledAttr} onchange="updateOfficer(${o.id},'gender',this.value)"><option value="">—</option><option ${o.gender==='ذكر'?'selected':''}>ذكر</option><option ${o.gender==='أنثى'?'selected':''}>أنثى</option></select></div>
        </div>
        <div class="row g-2">
          <div class="col-md-6"><label class="form-label">تاريخ الميلاد</label><input type="date" class="form-control" ${disabledAttr} value="${o.birth||''}" onchange="updateOfficer(${o.id},'birth',this.value)"></div>
          <div class="col-md-6"><label class="form-label">حالة</label><div class="d-flex align-items-center gap-2">${renderStatusBadge(o.status||'active')}${canManage?`<button class="btn btn-sm btn-outline-primary" onclick="setOfficerStatus(${o.id},'active')">نشط</button><button class="btn btn-sm btn-outline-warning" onclick="setOfficerStatus(${o.id},'retired')">تقاعد</button><button class="btn btn-sm btn-outline-secondary" onclick="setOfficerStatus(${o.id},'transferred')">نقل</button>`:''}</div></div>
        </div>
        <div class="row g-2">
           <div class="col-md-6"><label class="form-label">نوع التعيين</label><select id="off_${o.id}_off_admission" class="form-select" ${disabledAttr} onchange="setOfficerAdmissionType(${o.id}, this.value)"><option value="fresh" ${o.admissionType!=='transferred'?'selected':''}>خريج جديد</option><option value="transferred" ${o.admissionType==='transferred'?'selected':''}>منقول من جهة أخرى</option></select></div>
          <div class="col-md-6" id="off_${o.id}_off_hiring_fields" style="display:${transferInVisible?'none':'block'};">
            <label class="form-label">تاريخ التعيين</label>
            <input type="date" class="form-control" ${disabledAttr} value="${o.hiringDate||''}" onchange="updateOfficer(${o.id},'hiringDate',this.value)">
          </div>
          <div class="col-md-12" id="off_${o.id}_off_transfer_fields" style="display:${transferInVisible?'block':'none'};">
            <label class="form-label">جهة النقل من (الإدارات العامة ومديريات الأمن)</label>
            <select class="form-select mb-2" ${disabledAttr} onchange="updateOfficer(${o.id},'transferFromDeptId',+this.value||null)">
              <option value="">— اختر جهة —</option>
              ${transferDepts.map(d=>`<option value="${d.id}" ${d.id===o.transferFromDeptId?'selected':''}>${escapeHtml(d.name)}</option>`).join('')}
            </select>
            <label class="form-label">تاريخ النقل للداخل</label>
            <input type="date" class="form-control" ${disabledAttr} value="${o.transferDate||''}" onchange="updateOfficer(${o.id},'transferDate',this.value)">
          </div>
        </div>
        <div class="row g-2" id="off_${o.id}_off_transfer_out_fields" style="display:${transferOutVisible?'block':'none'};">
          <div class="col-md-12">
            <label class="form-label">جهة النقل إلى (الإدارات العامة ومديريات الأمن)</label>
            <select class="form-select mb-2" ${disabledAttr} onchange="updateOfficer(${o.id},'transferToDeptId',+this.value||null)">
              <option value="">— اختر جهة —</option>
              ${transferDepts.map(d=>`<option value="${d.id}" ${d.id===o.transferToDeptId?'selected':''}>${escapeHtml(d.name)}</option>`).join('')}
            </select>
            <label class="form-label">تاريخ النقل للخارج</label>
            <input type="date" class="form-control" ${disabledAttr} value="${o.transferToDate||''}" onchange="updateOfficer(${o.id},'transferToDate',this.value)">
          </div>
        </div>
        <div class="mb-2"><label class="form-label">العنوان</label><input class="form-control" ${disabledAttr} value="${escapeHtml(o.address||'')}" onchange="updateOfficer(${o.id},'address',this.value)"></div>
        <div class="row g-2">
          <div class="col-md-6"><label class="form-label">شخص طوارئ</label><input class="form-control" ${disabledAttr} value="${escapeHtml(o.emgName||'')}" onchange="updateOfficer(${o.id},'emgName',this.value)"></div>
          <div class="col-md-6"><label class="form-label">هاتف الطوارئ</label><input class="form-control" ${disabledAttr} value="${escapeHtml(o.emgPhone||'')}" onchange="updateOfficer(${o.id},'emgPhone',this.value)"></div>
        </div>
      </div></div>`;
    }
    function setOfficerStatus(id,status){
      const o = officers.find(x=>x.id===id);
      if(!o) return;
      if(status==='transferred'){
        o.status='transferred';
        o.transferNote='منقول خارج الإدارة';
        o.deptId = null;
        o.archivedAt = new Date().toISOString();
      }
      else if(status==='retired'){
        o.status='retired';
        o.archivedAt = new Date().toISOString();
      }
      else {
        o.status='active';
        o.archivedAt = null;
      }
      saveAll();
      renderTab('officers');
      safeShowToast('تم تحديث الحالة','success');
    }
    function addOfficerForm(){
      const area = document.getElementById('officerEditorArea');
      area.innerHTML = `<div class="card mt-2"><div class="card-header">ضابط جديد</div><div class="card-body">
        <div class="row g-2">
          <div class="col-md-4"><input id="new_off_name" class="form-control" placeholder="الاسم"></div>
          <div class="col-md-2"><select id="new_off_rank" class="form-select">${ranks.map(r=>`<option>${r}</option>`).join('')}</select></div>
          <div class="col-md-2">
            <div class="d-flex align-items-center gap-1">
              <input id="new_off_badge_1" class="form-control" placeholder="رقم 1" required style="min-width:0;">
              <span class="text-muted">/</span>
              <input id="new_off_badge_2" class="form-control" placeholder="رقم 2" required style="min-width:0;">
              <span class="text-muted">/</span>
              <input id="new_off_badge_3" class="form-control" placeholder="اختياري" style="min-width:0;">
            </div>
          </div>
           <div class="col-md-2"><select id="new_off_dept" class="form-select"><option value="">قسم</option>${getInternalDepartments().map(d=>`<option value="${d.id}">${escapeHtml(d.name)}</option>`).join('')}</select></div>
          <div class="col-md-2"><select id="new_off_job" class="form-select"><option value="">مسمى</option>${jobTitles.map(j=>`<option value="${j.id}">${escapeHtml(j.name)}</option>`).join('')}</select></div>
           <div class="col-md-2"><input id="new_off_phone" class="form-control" placeholder="هاتف"></div>
          <div class="col-md-2"><input id="new_off_gender" class="form-control" placeholder="الجنس"></div>
          <div class="col-md-2"><input id="new_off_birth" type="date" class="form-control" placeholder="ميلاد"></div>
          <div class="col-md-4"><input id="new_off_address" class="form-control" placeholder="العنوان"></div>
          <div class="col-md-3"><input id="new_off_emgName" class="form-control" placeholder="طوارئ: الاسم"></div>
          <div class="col-md-3"><input id="new_off_emgPhone" class="form-control" placeholder="طوارئ: الهاتف"></div>
           <div class="col-md-3"><select id="new_off_admission" class="form-select" onchange="toggleTransferFields('new')"><option value="fresh">خريج جديد</option><option value="transferred">منقول من جهة أخرى</option></select></div>
          <div class="col-md-3" id="new_off_hiring_fields">
            <input id="new_off_hiringDate" type="date" class="form-control" placeholder="تاريخ التعيين">
          </div>
          <div class="col-md-6" id="new_off_transfer_fields" style="display:none;">
            <div class="d-flex gap-2">
              <select id="new_off_transferDept" class="form-select"><option value="">— جهة النقل من —</option>${getTransferDepartments().map(d=>`<option value="${d.id}">${escapeHtml(d.name)}</option>`).join('')}</select>
              <input id="new_off_transferDate" type="date" class="form-control" placeholder="تاريخ النقل للداخل">
            </div>
          </div>
          </div>
        </div>
        <div class="text-end mt-2"><button class="btn btn-secondary" onclick="document.getElementById('officerEditorArea').innerHTML=''">إلغاء</button> <button class="btn btn-primary" onclick="createOfficer()">حفظ</button></div>
      </div></div>`;
      toggleTransferFields('new');
    }
    function createOfficer(){
      const name = document.getElementById('new_off_name').value.trim();
      const rank = document.getElementById('new_off_rank').value;
      if(!name) return showToast('ادخل الاسم','danger');
      const badgePart1 = document.getElementById('new_off_badge_1').value;
      const badgePart2 = document.getElementById('new_off_badge_2').value;
      const badgePart3 = document.getElementById('new_off_badge_3').value;
      if(!badgePart1.trim() || !badgePart2.trim()){
        return showToast('رقما الأقدمية الأول والثاني مطلوبان','danger');
      }
     const badgeValue = buildBadgeValue(
        badgePart1,
        badgePart2,
        badgePart3
      );
      if(isDuplicateOfficer(name, rank, badgeValue)){
        return showToast('هذا الضابط مسجل بالفعل بنفس الرتبة ورقم الأقدمية','warning');
      }
     const admissionType = document.getElementById('new_off_admission').value || 'fresh';
      const hiringDate = admissionType !== 'transferred' ? (document.getElementById('new_off_hiringDate').value || '') : '';
      const transferFromDeptId = admissionType === 'transferred' ? (+document.getElementById('new_off_transferDept').value || null) : null;
       const transferToDeptId = null;
      const transferDate = admissionType === 'transferred' ? (document.getElementById('new_off_transferDate').value || '') : '';
      const transferToDate = '';
      officers.push({id, name, rank, badge: badgeValue, deptId: +document.getElementById('new_off_dept').value || null, jobTitleId: +document.getElementById('new_off_job').value || null, subs: [], status:'active', transferNote:'', phone:document.getElementById('new_off_phone').value, gender:document.getElementById('new_off_gender').value, birth:document.getElementById('new_off_birth').value, address:document.getElementById('new_off_address').value, emgName:document.getElementById('new_off_emgName').value, emgPhone:document.getElementById('new_off_emgPhone').value, admissionType, hiringDate, transferFromDeptId, transferToDeptId, transferDate, transferToDate, archivedAt: null});
      saveAll();
      document.getElementById('officerEditorArea').innerHTML='';
      SELECTED_OFFICER_ID = id;
      logActivity('إضافة ضابط', {target: name, rank});
      renderTab('officers');
    }
    function updateOfficerBadgeFromInputs(id){
      const part1 = document.getElementById(`badge_${id}_1`)?.value || '';
      const part2 = document.getElementById(`badge_${id}_2`)?.value || '';
      const part3 = document.getElementById(`badge_${id}_3`)?.value || '';
      if(!part1.trim() || !part2.trim()){
        safeShowToast('رقما الأقدمية الأول والثاني مطلوبان','warning');
        return;
      }
      updateOfficer(id, 'badge', buildBadgeValue(part1, part2, part3));
    }
    function updateOfficer(id,key,val){ const o = officers.find(x=>x.id===id); if(!o) return; o[key]=val; saveAll(); }
   function getExternalDepartments(){
      return departments.filter(d=> (d.groupKey || 'internal') === 'external');
    }
    function getTransferDepartments(){
      return departments.filter(d=> (d.groupKey || 'internal') === 'transfer');
    }
    function getInternalDepartments(){
      return departments.filter(d=> (d.groupKey || 'internal') !== 'transfer');
    }
     function toggleTransferFields(prefix){
      const type = document.getElementById(`${prefix}_off_admission`)?.value;
      const transferContainer = document.getElementById(`${prefix}_off_transfer_fields`);
      if(transferContainer) transferContainer.style.display = type === 'transferred' ? 'block' : 'none';
      const hiringContainer = document.getElementById(`${prefix}_off_hiring_fields`);
      if(hiringContainer) hiringContainer.style.display = type === 'transferred' ? 'none' : 'block';
 }
    function setOfficerAdmissionType(id, val){
      const o = officers.find(x=>x.id===id);
      if(!o) return;
      o.admissionType = val;
     if(val !== 'transferred'){
        o.transferFromDeptId = null;
        o.transferDate = '';
      } else {
        o.hiringDate = '';
      }
      saveAll();
      renderTab('officers');
    }
    function deleteOfficer(id){ if(!confirm('حذف الضابط؟')) return; const removed = officers.find(o=>o.id===id); officers = officers.filter(o=>o.id!==id); delete officerLimits[id]; saveAll(); logActivity('حذف ضابط', {target: removed?.name || id}); renderTab('officers'); }
    function openSubordinatesForm(id){
      const o = officers.find(x=>x.id===id); if(!o) return;
      const area = document.getElementById('officerEditorArea');
      const options = officers.filter(x=>x.id!==id).map(x=>`<option value="${x.id}" ${(o.subs||[]).includes(x.id)?'selected':''}>${escapeHtml(x.name)}</option>`).join('');
      area.innerHTML = `<div class="card mt-2"><div class="card-header">توابع ${escapeHtml(o.name)}</div><div class="card-body">
        <select id="subs_select" multiple class="form-select" style="min-height:140px">${options}</select>
        <div class="text-end mt-2"><button class="btn btn-secondary" onclick="document.getElementById('officerEditorArea').innerHTML=''">إلغاء</button> <button class="btn btn-primary" onclick="saveSubs(${id})">حفظ</button></div>
      </div></div>`;
    }
    function saveSubs(id){ const o = officers.find(x=>x.id===id); if(!o) return; o.subs = Array.from(document.getElementById('subs_select').selectedOptions||[]).map(op=>+op.value); saveAll(); renderTab('officers'); }
       function exportOfficersCSV(){
      const header = 'id,name,rank,badge,deptId,jobTitleId,phone,gender,birth,address,emgName,emgPhone,status,admissionType,hiringDate,transferFromDeptId,transferToDeptId,transferDate,transferToDate\n';
      const lines = officers.map(o=>[
        o.id,o.name,o.rank,o.badge||'',o.deptId||'',o.jobTitleId||'',
        o.phone||'',o.gender||'',o.birth||'',o.address||'',o.emgName||'',o.emgPhone||'',
        o.status||'',o.admissionType||'fresh',o.hiringDate||'',o.transferFromDeptId||'',o.transferToDeptId||'',o.transferDate||'',o.transferToDate||''
      ].map(csvEscape).join(','));
      downloadCSVFile('officers.csv', header + lines.join('\n'));
    }
    function exportOfficersFullCSV(){
      const header = 'id,name,rank,badge,status,deptId,deptName,jobTitleId,jobTitle,subs,phone,gender,birth,address,emgName,emgPhone,admissionType,hiringDate,transferFromDeptId,transferFromDeptName,transferToDeptId,transferToDeptName,transferDate,transferToDate,transferNote\n';
      const lines = officers.map(o=>{
        const deptName = o.deptId ? (departments.find(d=>d.id===o.deptId)?.name || '') : '';
        const jobTitle = o.jobTitleId ? (jobTitles.find(j=>j.id===o.jobTitleId)?.name || '') : '';
        const transferFromName = o.transferFromDeptId ? (departments.find(d=>d.id===o.transferFromDeptId)?.name || '') : '';
        const transferToName = o.transferToDeptId ? (departments.find(d=>d.id===o.transferToDeptId)?.name || '') : '';
        return [
          o.id,o.name,o.rank,o.badge||'',o.status||'',
          o.deptId||'',deptName,
          o.jobTitleId||'',jobTitle,
          (o.subs||[]).join('|'),
          o.phone||'',o.gender||'',o.birth||'',o.address||'',o.emgName||'',o.emgPhone||'',
          o.admissionType||'fresh',o.hiringDate||'',
          o.transferFromDeptId||'',transferFromName,
          o.transferToDeptId||'',transferToName,
          o.transferDate||'',o.transferToDate||'',
          o.transferNote||''
        ].map(csvEscape).join(',');
      });
      downloadCSVFile('officers_full.csv', header + lines.join('\n'));
    }
    function downloadOfficersTemplateCSV(){ downloadCSVFile('officers_template.csv','id,name,rank,badge,deptId,jobTitleId,phone,gender,birth,address,emgName,emgPhone,status,admissionType,hiringDate,transferFromDeptId,transferToDeptId,transferDate,transferToDate'); }
    function importOfficersCSV(input){
      const f = input.files && input.files[0]; if(!f) return;
      const reader = new FileReader(); reader.onload = ()=>{
        const rows = reader.result.split(/\r?\n/).slice(1).filter(Boolean);
         let added = 0;
        let skipped = 0;
        let merged = 0;
        rows.forEach(line=>{
           const [id,name,rank,badge,deptId,jobTitleId,phone,gender,birth,address,emgName,emgPhone,status,admissionType,hiringDate,transferFromDeptId,transferToDeptId,transferDate,transferToDate] = line.split(',');
          if(!name) return;
          const cleanName = name.trim();
          const cleanRank = (rank || ranks[0]).trim();
          const cleanBadge = (badge || '').trim();
          if(isDuplicateOfficer(cleanName, cleanRank, cleanBadge)){
            const existing = officers.find(o=>normalizeOfficerKey(o.name, o.rank, o.badge) === normalizeOfficerKey(cleanName, cleanRank, cleanBadge));
            if(existing){
              const mergedFields = {
                deptId: deptId? +deptId : null,
                jobTitleId: jobTitleId? +jobTitleId : null,
                phone: phone || '',
                gender: gender || '',
                birth: birth || '',
                address: address || '',
                emgName: emgName || '',
                emgPhone: emgPhone || '',
                status: status || '',
                admissionType: admissionType || '',
                hiringDate: hiringDate || '',
                transferFromDeptId: transferFromDeptId ? +transferFromDeptId : null,
                transferToDeptId: transferToDeptId ? +transferToDeptId : null,
                transferDate: transferDate || '',
                transferToDate: transferToDate || ''
              };
              const fieldsToUpdate = [
                ['deptId', v=> v != null],
                ['jobTitleId', v=> v != null],
                ['phone', v=> v],
                ['gender', v=> v],
                ['birth', v=> v],
                ['address', v=> v],
                ['emgName', v=> v],
                ['emgPhone', v=> v],
                ['status', v=> v],
                ['admissionType', v=> v],
                ['hiringDate', v=> v],
                ['transferFromDeptId', v=> v != null],
                ['transferToDeptId', v=> v != null],
                ['transferDate', v=> v],
                ['transferToDate', v=> v]
              ];
              let updated = false;
              fieldsToUpdate.forEach(([key, hasValue])=>{
                if((existing[key] == null || existing[key] === '') && hasValue(mergedFields[key])){
                  existing[key] = mergedFields[key];
                  updated = true;
                }
              });
              if(updated) merged++;
            } else {
              skipped++;
            }
            return;
          }
          officers.push({
            id: id? +id : Date.now()+Math.random(),
            name: cleanName,
            rank: cleanRank,
            badge: cleanBadge,
            deptId: deptId? +deptId : null,
            jobTitleId: jobTitleId? +jobTitleId : null,
            subs: [],
            phone: phone || '',
            gender: gender || '',
            birth: birth || '',
            address: address || '',
            emgName: emgName || '',
            emgPhone: emgPhone || '',
            status: status || 'active',
            admissionType: admissionType || 'fresh',
            hiringDate: hiringDate || '',
           transferFromDeptId: transferFromDeptId ? +transferFromDeptId : null,
            transferToDeptId: transferToDeptId ? +transferToDeptId : null,
           transferDate: transferDate || '',
            transferToDate: transferToDate || ''
          });
          added++;
        });
        saveAll(); renderTab('officers');
        if(merged){
          showToast(`تم دمج بيانات ${merged} ضابط مكرر`, 'info');
        }
        if(skipped){
          showToast(`تم تجاهل ${skipped} ضابط مكرر`, 'warning');
        } else if(added){
          showToast('تم استيراد الضباط', 'success');
        }
      };
      reader.readAsText(f);
    }
   /* ========= Departments ========= */
    function renderDepartmentsTab(){
      if(SETTINGS.authEnabled && (!CURRENT_USER || currentRole()==='editor')) return signInInlineFormMarkup() + '<div class="alert alert-warning">ليست لديك صلاحية تعديل الأقسام.</div>';
      const internalLabel = (SETTINGS.deptGroups||[]).find(g=>g.key==='internal')?.name || 'الأقسام الداخلية';
      const externalLabel = (SETTINGS.deptGroups||[]).find(g=>g.key==='external')?.name || 'الأقسام الخارجية';
      const internalTree = renderDepartmentGroupTree(null,0,'internal', false, 'departments');
      const externalTree = renderDepartmentGroupTree(null,0,'external', false, 'departments');
      return `<div class="card"><div class="card-header d-flex justify-content-between"><div>الأقسام (داخل ديوان الإدارة العامة وخارجها)</div><div class="d-flex gap-2"><button class="btn btn-success btn-sm" onclick="openDepartmentForm('add',null,null,'internal','internal','departments')">إضافة قسم</button><button class="btn btn-outline-primary btn-sm" onclick="openDepartmentForm('add',null,null,'external','external','departments')">إضافة قسم خارجي</button></div></div>
        <div class="card-body">
          <div class="d-flex flex-wrap gap-2 mb-3">
            <button class="btn btn-outline-secondary btn-sm" onclick="downloadDepartmentsTemplateCSV()">تحميل قالب CSV للأقسام</button>
            <label class="btn btn-outline-primary btn-sm mb-0">استيراد CSV للأقسام <input type="file" id="departmentsCsvInput" accept=".csv" style="display:none" onchange="importDepartmentsCSV(this)"></label>
          </div>
           <div class="row g-3">
            <div class="col-md-6">
              <div class="d-flex justify-content-between align-items-center mb-2">
                <div class="fw-bold">${escapeHtml(internalLabel)}</div>
                <div class="d-flex gap-1">
                  <button class="btn btn-outline-secondary btn-sm" onclick="setDeptGroupCollapse('internal', true, 'departments')">طي الكل</button>
                  <button class="btn btn-outline-secondary btn-sm" onclick="setDeptGroupCollapse('internal', false, 'departments')">توسيع الكل</button>
                </div>
              </div>
              ${internalTree}
            </div>
            <div class="col-md-6">
              <div class="d-flex justify-content-between align-items-center mb-2">
                <div class="fw-bold">${escapeHtml(externalLabel)}</div>
                <div class="d-flex gap-1">
                  <button class="btn btn-outline-secondary btn-sm" onclick="setDeptGroupCollapse('external', true, 'departments')">طي الكل</button>
                  <button class="btn btn-outline-secondary btn-sm" onclick="setDeptGroupCollapse('external', false, 'departments')">توسيع الكل</button>
                </div>
              </div>
              ${externalTree}
            </div>
          </div>
          <div class="small-muted mt-2">المناطق الجغرافية تتبع الإدارة العامة لكنها تقع في خارج ديوان الإدارة.</div>
          <div class="mt-3"><div id="departmentFormArea"></div></div>
        </div></div>`;
    }
    function renderTransferDepartmentsTab(){
      if(SETTINGS.authEnabled && (!CURRENT_USER || currentRole()==='editor')) return signInInlineFormMarkup() + '<div class="alert alert-warning">ليست لديك صلاحية تعديل جهات النقل.</div>';
      const transferLabel = (SETTINGS.deptGroups||[]).find(g=>g.key==='transfer')?.name || 'جهات النقل';
      const transferTree = renderDepartmentGroupTree(null,0,'transfer', false, 'transfer-departments');
      return `<div class="card"><div class="card-header d-flex justify-content-between"><div>${escapeHtml(transferLabel)}</div><div class="d-flex gap-2"><button class="btn btn-success btn-sm" onclick="openDepartmentForm('add',null,null,'transfer','transfer','transfer-departments')">إضافة جهة نقل</button></div></div>
        <div class="card-body">
          <div class="d-flex flex-wrap gap-2 mb-3">
            <button class="btn btn-outline-secondary btn-sm" onclick="downloadTransferDepartmentsTemplateCSV()">تحميل قالب CSV لجهات النقل</button>
            <label class="btn btn-outline-primary btn-sm mb-0">استيراد CSV لجهات النقل <input type="file" id="transferDepartmentsCsvInput" accept=".csv" style="display:none" onchange="importTransferDepartmentsCSV(this)"></label>
          </div>
          <div class="row g-3">
            <div class="col-md-6">
              <div class="d-flex justify-content-between align-items-center mb-2">
                <div class="fw-bold">قائمة الجهات</div>
                <div class="d-flex gap-1">
                  <button class="btn btn-outline-secondary btn-sm" onclick="setDeptGroupCollapse('transfer', true, 'transfer-departments')">طي الكل</button>
                  <button class="btn btn-outline-secondary btn-sm" onclick="setDeptGroupCollapse('transfer', false, 'transfer-departments')">توسيع الكل</button>
                </div>
              </div>
              ${transferTree}
            </div>
            <div class="col-md-6"><div id="departmentFormArea"></div></div>
          </div>
          <div class="small-muted mt-2">جهات النقل هي الإدارات العامة ومديريات الأمن، ويتم إدخالها بالاسم فقط ويمكن رفعها عبر ملف CSV.</div>
        </div></div>`;
    }
    function downloadDepartmentsTemplateCSV(){
      downloadCSVFile('departments_template.csv','id,name,groupKey,parentId,headId,upperTitle,upperOfficerId');
    }
    function downloadTransferDepartmentsTemplateCSV(){
      downloadCSVFile('transfer_departments_template.csv','name');
    }
   function importDepartmentsCSV(input){
      const f = input.files && input.files[0]; if(!f) return;
      const reader = new FileReader(); reader.onload = ()=>{
        const rows = reader.result.split(/\r?\n/).slice(1).filter(Boolean);
        rows.forEach(line=>{
          const [id,name,groupKey,parentId,headId,upperTitle,upperOfficerId] = line.split(',');
          if(!name) return;
          departments.push({
            id: id? +id : Date.now()+Math.random(),
            name: name.trim(),
            groupKey: (groupKey||'internal').trim() || 'internal',
            parentId: parentId? +parentId : null,
            headId: headId? +headId : null,
            upperTitle: upperTitle || '',
            upperOfficerId: upperOfficerId? +upperOfficerId : null
          });
        });
        saveAll();
        renderTab('departments');
      };
      reader.readAsText(f);
    }
     function importTransferDepartmentsCSV(input){
      const f = input.files && input.files[0]; if(!f) return;
      const reader = new FileReader(); reader.onload = ()=>{
        const rows = reader.result.split(/\r?\n/).slice(1).filter(Boolean);
        let added = 0;
        let skipped = 0;
        rows.forEach(line=>{
          const [name] = line.split(',');
          if(!name) return;
          if(isDuplicateTransferDepartment(name.trim())){
            skipped++;
            return;
          }
          departments.push({
            id: Date.now() + Math.random(),
            name: name.trim(),
            groupKey: 'transfer',
            parentId: null,
            headId: null,
            upperTitle: '',
            upperOfficerId: null
          });
          added++;
        });
        saveAll();
        renderTab('transfer-departments');
        if(skipped){
          showToast(`تم تجاهل ${skipped} جهة مكررة`, 'warning');
        } else if(added){
          showToast('تم استيراد جهات النقل', 'success');
        }
      };
      reader.readAsText(f);
    }
     function renderDepartmentGroupTree(parentId,level,groupKey,allowDrag=false,returnTab='departments'){
      const children = departments.filter(d=> (d.parentId||null) === (parentId||null) && d.groupKey===groupKey);
      if(!children.length) return parentId ? '' : `<p class="text-muted py-3">لا توجد أقسام</p>`;
      return `<ul class="list-unstyled ms-${level*3}">${children.map(d=>{
        const hasChildren = departments.some(ch=> (ch.parentId||null) === d.id && ch.groupKey===groupKey);
        const collapsed = !!deptCollapseState[d.id];
        const childTree = hasChildren ? renderDepartmentGroupTree(d.id,level+1,groupKey,allowDrag,returnTab) : '';
        return `<li class="mb-2 p-2 border rounded" ${allowDrag ? 'draggable="true" ondragstart="onDeptDragStart(event,'+d.id+')"' : ''}>
          <div class="d-flex justify-content-between align-items-start">
            <div>
              <div class="d-flex align-items-center gap-2">
                ${hasChildren ? `<button class="btn btn-sm btn-outline-secondary" onclick="toggleDeptCollapse(${d.id},'${returnTab}')">${collapsed ? 'عرض' : 'إخفاء'}</button>` : ''}
                <strong>${escapeHtml(d.name)}</strong> ${renderDeptGroupBadge(d.groupKey)}
              </div>
              ${d.headId? '<div class="small text-muted">رئيس: '+escapeHtml((officers.find(o=>o.id===d.headId)||{name:""}).name)+'</div>':''}
              ${d.upperTitle? '<div class="small text-muted">مسمى إدارة عليا: '+escapeHtml(d.upperTitle) + (d.upperOfficerId? ' — '+escapeHtml((officers.find(o=>o.id===d.upperOfficerId)||{name:""}).name) : '') + '</div>':''}
            </div>
            <div>
              <button class="btn btn-sm btn-outline-primary me-1" onclick="openDepartmentForm('edit',${d.id},null,'${groupKey}','${groupKey}','${returnTab}')">تعديل</button>
              <button class="btn btn-sm btn-outline-success me-1" onclick="openDepartmentForm('add',null,${d.id},'${groupKey}','${groupKey}','${returnTab}')">وحدة فرعية</button>
              <button class="btn btn-sm btn-outline-danger me-1" onclick="markDepartmentDelete(${d.id},'${returnTab}')">حذف</button>
            </div>
          </div>
          ${hasChildren ? `<div style="margin-top:8px;display:${collapsed?'none':'block'};">${childTree}</div>` : ''}
        </li>`;
      }).join('')}</ul>`;
    }
    function renderDeptGroupBadge(groupKey){
      const name = (SETTINGS.deptGroups||[]).find(g=>g.key===groupKey)?.name || '';
      const cls = groupKey==='external'?'status-transferred':'status-active';
      return name ? `<span class="badge-status ${cls}" style="margin-inline-start:6px;">${escapeHtml(name)}</span>` : '';
    }
    function toggleDeptCollapse(id, returnTab='departments'){
      deptCollapseState[id] = !deptCollapseState[id];
      renderTab(returnTab);
    }
    function setDeptGroupCollapse(groupKey, collapsed, returnTab='departments'){
      departments.filter(d=>d.groupKey===groupKey).forEach(d=>{ deptCollapseState[d.id] = collapsed; });
      renderTab(returnTab);
    }
    function onDeptDragStart(ev, id){
      ev.dataTransfer.setData('text/plain', id);
    }
    function allowDeptDrop(ev){
      ev.preventDefault();
    }
    function handleDeptDrop(ev, targetGroup){
      ev.preventDefault();
      const id = +ev.dataTransfer.getData('text/plain');
      const d = departments.find(x=>x.id===id);
      if(!d || !targetGroup) return;
      d.groupKey = targetGroup;
      saveAll();
      renderTab('departments');
      showToast('تم نقل القسم للمجموعة الجديدة','success');
    }
    function updateDeptGroupName(key, val){
      const name = (val||'').trim();
      if(!name) return showToast('الاسم مطلوب','danger');
      SETTINGS.deptGroups = (SETTINGS.deptGroups||defaultSettings.deptGroups.slice()).map(g=> g.key===key ? Object.assign({}, g, {name}) : g);
      saveAll();
      renderTab('departments');
      showToast('تم تحديث أسماء المجموعات','success');
    }
    function openDepartmentForm(mode,id=null,parentId=null,forcedGroup=null,scopeGroup=null,returnTab='departments'){
      const area = document.getElementById('departmentFormArea');
      let name='', headId='', pid = parentId || '', upperTitle='', upperOfficerId='', groupKey='internal';
       if(mode==='edit' && id){ const d = departments.find(x=>x.id===id); if(d){ name=d.name; headId=d.headId||''; pid=d.parentId||''; upperTitle = d.upperTitle || ''; upperOfficerId = d.upperOfficerId || ''; groupKey = d.groupKey || 'internal'; } }
      else if(mode==='add'){ pid = parentId || ''; groupKey = forcedGroup || 'internal'; }
      if(scopeGroup) groupKey = scopeGroup;
      const idValue = id ?? 'null';
      const isTransferGroup = groupKey === 'transfer';
      const parentOptions = (scopeGroup ? departments.filter(d=> (d.groupKey||'internal')===scopeGroup) : departments)
        .map(d=>`<option value="${d.id}" ${d.id==pid?'selected':''}>${escapeHtml(d.name)}</option>`).join('');
      const groupSelector = scopeGroup ? '' : `<div class="mb-2"><label class="form-label">مجموعة القسم</label><select id="dept_group" class="form-select">${(SETTINGS.deptGroups||[]).map(g=>`<option value="${g.key}" ${g.key===groupKey?'selected':''}>${escapeHtml(g.name)}</option>`).join('')}</select></div>`;
      const transferNote = isTransferGroup ? '<div class="alert alert-info mb-2">جهات النقل تُسجّل بالاسم فقط.</div>' : '';
      const extraFields = isTransferGroup ? '' : `
        <div class="mb-2"><label class="form-label">تابع لـ</label><select id="dept_parent" class="form-select"><option value="">— لا شيء —</option>${parentOptions}</select></div>
        <div class="mb-2"><label class="form-label">رئيس القسم</label><select id="dept_head" class="form-select"><option value="">— لا شيء —</option>${officers.map(o=>`<option value="${o.id}" ${o.id==headId?'selected':''}>${escapeHtml(o.rank)} ${escapeHtml(o.name)}</option>`).join('')}</select></div>
        <div class="mb-2"><label class="form-label">مسمى الإدارة العليا (اختياري)</label><input id="dept_upperTitle" class="form-control" value="${escapeHtml(upperTitle)}"></div>
        <div class="mb-2"><label class="form-label">ضابط الإدارة العليا (اختياري)</label><select id="dept_upperOfficer" class="form-select"><option value="">— لا شيء —</option>${officers.map(o=>`<option value="${o.id}" ${o.id==upperOfficerId?'selected':''}>${escapeHtml(o.rank)} ${escapeHtml(o.name)}</option>`).join('')}</select></div>`;
      area.innerHTML = `<div class="card"><div class="card-header">${mode==='add'?'إضافة قسم':'تعديل قسم'}</div><div class="card-body">
        ${transferNote}
        <div class="mb-2"><label class="form-label">اسم</label><input id="dept_name" class="form-control" value="${escapeHtml(name)}" onkeydown="handleDepartmentFormKey(event,'${mode}',${idValue},'${groupKey}','${returnTab}')"></div>
        ${extraFields}
        ${groupSelector}
        <div class="text-end"><button class="btn btn-secondary" onclick="document.getElementById('departmentFormArea').innerHTML=''">إلغاء</button> <button class="btn btn-primary" onclick="saveDepartment('${mode}',${idValue},'${groupKey}','${returnTab}')">حفظ</button></div>
      </div></div>`;
    }
    function handleDepartmentFormKey(event, mode, id, groupKey, returnTab){
      if(event.key !== 'Enter') return;
      event.preventDefault();
      saveDepartment(mode, id, groupKey, returnTab);
    }
    function isDuplicateTransferDepartment(name, ignoreId=null){
      const target = (name || '').trim().toLowerCase();
      if(!target) return false;
      return departments.some(d => d.groupKey === 'transfer' && d.id !== ignoreId && (d.name || '').trim().toLowerCase() === target);
    }
    function saveDepartment(mode,id,forcedGroup=null,returnTab='departments'){
      const name = document.getElementById('dept_name').value.trim();
      const parentEl = document.getElementById('dept_parent');
      const headEl = document.getElementById('dept_head');
      const upperTitleEl = document.getElementById('dept_upperTitle');
      const upperOfficerEl = document.getElementById('dept_upperOfficer');
      const parentId = parentEl && parentEl.value ? +parentEl.value : null;
      const headId = headEl && headEl.value ? +headEl.value : null;
      const upperTitle = upperTitleEl ? upperTitleEl.value.trim() : '';
      const upperOfficerId = upperOfficerEl && upperOfficerEl.value ? +upperOfficerEl.value : null;
      const groupKey = forcedGroup || (document.getElementById('dept_group') ? document.getElementById('dept_group').value : 'internal');
      if(!name) return showToast('أدخل اسم القسم','danger');
      if(groupKey === 'transfer' && isDuplicateTransferDepartment(name, mode === 'edit' ? id : null)){
        return showToast('الجهة موجودة بالفعل','warning');
      }
      if(mode==='add'){ departments.push({id:Date.now(),name,parentId,headId,upperTitle,upperOfficerId,groupKey}); }
      else if(mode==='edit'){ const d = departments.find(x=>x.id===id); if(d){ d.name=name; d.parentId=parentId; d.headId=headId; d.upperTitle=upperTitle; d.upperOfficerId=upperOfficerId; d.groupKey=groupKey; } }
      saveAll(); document.getElementById('departmentFormArea').innerHTML=''; renderTab(returnTab);
    }
    function markDepartmentDelete(id,returnTab='departments'){
      const area = document.getElementById('departmentFormArea');
      area.innerHTML = `<div class="alert alert-danger">هل تريد حذف هذا القسم وجميع وحداته الفرعية؟<div class="text-end mt-2"><button class="btn btn-secondary" onclick="document.getElementById('departmentFormArea').innerHTML=''">إلغاء</button> <button class="btn btn-danger" onclick="deleteDepartmentRecursive(${id},'${returnTab}')">تأكيد حذف</button></div></div>`;
    }
    function deleteDepartmentRecursive(id,returnTab='departments'){
      const toDelete=[id];
      for(let i=0;i<toDelete.length;i++){ const cur=toDelete[i]; departments.filter(dd=>dd.parentId===cur).forEach(ch=>toDelete.push(ch.id)); }
      departments = departments.filter(d=>!toDelete.includes(d.id));
      officers.forEach(o=>{ if(toDelete.includes(o.deptId)) o.deptId = null; });
      saveAll(); document.getElementById('departmentFormArea').innerHTML=''; renderTab(returnTab);
    }
    /* ========= Job Titles ========= */
    function renderJobTitlesTab(){
      if(SETTINGS.authEnabled && (!CURRENT_USER || currentRole()==='editor')) return signInInlineFormMarkup() + '<div class="alert alert-warning">ليست لديك صلاحية تعديل المسميات.</div>';
      return `<div class="card"><div class="card-header d-flex justify-content-between"><div>المسميات الوظيفية</div><div><button class="btn btn-success btn-sm" onclick="openJobTitleForm('add')">إضافة مسمى</button></div></div>
        <div class="card-body">
          <div class="row"><div class="col-md-6">${renderJobTitleTree(null,0)}</div><div class="col-md-6"><div id="jobTitleFormArea"></div></div></div>
        </div></div>`;
    }
    function renderJobTitleTree(parentId, level){
      const children = jobTitles.filter(j=> (j.parentId||null) === (parentId||null));
      if(!children.length && !parentId) return `<p class="text-muted py-3">لا توجد مسميات وظيفية</p>`;
      return `<ul class="list-unstyled ms-${level*3}">${children.map(j=>`<li class="mb-2 p-2 border rounded">
        <div class="d-flex justify-content-between align-items-center">
          <div><strong>${escapeHtml(j.name)}</strong> ${j.isUpper?'<span class="small text-warning">[إدارة عليا]</span>':''}</div>
          <div>
            <button class="btn btn-sm btn-outline-primary me-1" onclick="openJobTitleForm('edit',${j.id})">تعديل</button>
            <button class="btn btn-sm btn-outline-success me-1" onclick="openJobTitleForm('add',null,${j.id})">فرعي</button>
            <button class="btn btn-sm btn-outline-danger" onclick="if(confirm('حذف المسمى ووحداته الفرعية؟')) deleteJobTitleRecursive(${j.id})">حذف</button>
          </div>
        </div>
        ${renderJobTitleTree(j.id, level+1)}
      </li>`).join('')}</ul>`;
    }
    function openJobTitleForm(mode,id=null,parentId=null){
      const area = document.getElementById('jobTitleFormArea');
      let name='', isUpper=false;
      if(mode==='edit' && id){
        const jt = jobTitles.find(x=>x.id===id); if(jt){ name=jt.name; isUpper=jt.isUpper; parentId=jt.parentId||null; }
      }
      area.innerHTML = `<div class="card"><div class="card-header">${mode==='add'?'إضافة مسمى':'تعديل مسمى'}</div><div class="card-body">
        <div class="mb-2"><label class="form-label">الاسم</label><input id="jt_name" class="form-control" value="${escapeHtml(name)}"></div>
        <div class="form-check mb-2"><input id="jt_isUpper" class="form-check-input" type="checkbox" ${isUpper?'checked':''}><label class="form-check-label">إدارة عليا</label></div>
        <div class="mb-2"><label class="form-label">تابع لـ</label><select id="jt_parent" class="form-select"><option value="">— لا شيء —</option>${jobTitles.map(j=>`<option value="${j.id}" ${j.id==parentId?'selected':''}>${escapeHtml(j.name)}</option>`).join('')}</select></div>
        <div class="text-end"><button class="btn btn-secondary" onclick="document.getElementById('jobTitleFormArea').innerHTML=''">إلغاء</button> <button class="btn btn-primary" onclick="saveJobTitle('${mode}',${id||''})">حفظ</button></div>
      </div></div>`;
    }
    function saveJobTitle(mode,id){
      const name = document.getElementById('jt_name').value.trim();
      const isUpper = !!document.getElementById('jt_isUpper').checked;
      const parentId = document.getElementById('jt_parent').value ? +document.getElementById('jt_parent').value : null;
      if(!name) return showToast('أدخل اسم المسمى','danger');
      if(mode==='add') jobTitles.push({id:Date.now(),name,parentId,isUpper});
      else if(mode==='edit'){ const jt = jobTitles.find(x=>x.id===id); if(jt){ jt.name=name; jt.isUpper=isUpper; jt.parentId=parentId; } }
      saveAll(); document.getElementById('jobTitleFormArea').innerHTML=''; renderTab('jobtitles');
    }
    function deleteJobTitleRecursive(id){
      const toDelete=[id];
      for(let i=0;i<toDelete.length;i++){ const cur=toDelete[i]; jobTitles.filter(j=>j.parentId===cur).forEach(ch=>toDelete.push(ch.id)); }
      jobTitles = jobTitles.filter(j=>!toDelete.includes(j.id));
      departments.forEach(d=>{ if(toDelete.includes(d.jobTitleId)) d.jobTitleId = null; });
      officers.forEach(o=>{ if(toDelete.includes(o.jobTitleId)) o.jobTitleId = null; });
      duties.forEach(d=>{ d.signingJobTitleIds = (d.signingJobTitleIds||[]).filter(id=>!toDelete.includes(id)); });
      saveAll(); renderTab('jobtitles');
    }

    /* ========= Officer Limits ========= */
    function renderOfficerLimitsTab(){
      if(SETTINGS.authEnabled && (!CURRENT_USER || currentRole()==='editor')) return signInInlineFormMarkup() + '<div class="alert alert-warning">لا تملك صلاحية تعديل الحدود.</div>';
      if(!officers.length) return '<div class="alert alert-info">أضف ضباطاً أولاً لضبط الحدود.</div>';
      if(!duties.length) return '<div class="alert alert-info">أضف أنواع الخدمات أولاً لضبط الحدود.</div>';
      const dutyHeaders = duties.map(d=>`<th>${escapeHtml(d.name)}</th>`).join('');
      const filteredOfficers = officers.filter(o=>{
        const deptGroup = (departments.find(d=>d.id===o.deptId)?.groupKey) || 'internal';
        if(officerLimitFilters.hideInternal && deptGroup==='internal') return false;
        if(officerLimitFilters.hideExternal && deptGroup==='external') return false;
        if(officerLimitFilters.hiddenRanks.includes(o.rank)) return false;
        return true;
      });
      const rows = filteredOfficers.map(o=>{
        const dutyInputs = duties.map(d=>{
          const v = getOfficerDutyLimit(o.id, d.id);
          return `<td><input type="number" min="0" class="form-control form-control-sm" value="${v==null?'':v}" onchange="setOfficerDutyLimit(${o.id},${d.id},this.value)"></td>`;
        }).join('');
        const totalVal = getOfficerTotalLimit(o.id);
        return `<tr><td>${escapeHtml(o.rank)}</td><td>${escapeHtml(o.name)}</td>${dutyInputs}<td><input type="number" min="0" class="form-control form-control-sm" value="${totalVal==null?'':totalVal}" onchange="setOfficerTotalLimit(${o.id},this.value)"></td></tr>`;
      }).join('');
      const rankToggles = ranks.map(r=>`<label class="form-check me-2"><input type="checkbox" ${officerLimitFilters.hiddenRanks.includes(r)?'checked':''} onchange="toggleLimitRankFilter('${r}', this.checked)"> إخفاء ${escapeHtml(r)}</label>`).join('');
      const visibleRanks = ranks.filter(r=>!officerLimitFilters.hiddenRanks.includes(r));
      const rankLimitTargets = `<option value="total">الإجمالي الشهري</option>` + duties.map(d=>`<option value="${d.id}">${escapeHtml(d.name)}</option>`).join('');
      const rankLimitOptions = visibleRanks.map(r=>`<option value="${escapeHtml(r)}">${escapeHtml(r)}</option>`).join('');
      const rankLimitExcluded = officerLimitFilters.rankLimitExcluded || [];
      const rankLimitExcludedOptions = officers.map(o=>`<option value="${o.id}" ${rankLimitExcluded.includes(o.id)?'selected':''}>${escapeHtml(o.rank)} ${escapeHtml(o.name)}</option>`).join('');
      const deptFilters = `
        <label class="form-check me-3"><input type="checkbox" ${officerLimitFilters.hideInternal?'checked':''} onchange="setLimitDeptVisibility('internal', this.checked)"> إخفاء أقسام داخلية</label>
        <label class="form-check"><input type="checkbox" ${officerLimitFilters.hideExternal?'checked':''} onchange="setLimitDeptVisibility('external', this.checked)"> إخفاء أقسام خارجية</label>`;
      return `<div class="card"><div class="card-header d-flex justify-content-between align-items-center"><div>حدود الخدمات لكل ضابط</div><div class="small-muted">تُطبق أثناء التوزيع الآلي والقوائم المنسدلة</div></div><div class="card-body">`+
             `<div class="alert alert-info"><div class="fw-bold mb-1">خيارات إخفاء من القائمة</div><div class="d-flex flex-wrap gap-2 align-items-center">${rankToggles || '<span class="text-muted">لا توجد رتب</span>'} ${deptFilters}</div></div>`+
             `<div class="alert alert-warning"><div class="fw-bold mb-2">تطبيق حد لرتبة</div>
              <div class="row g-2 align-items-end">
                <div class="col-md-3"><label class="form-label">الرتبة</label><select id="rank_limit_rank" class="form-select">${rankLimitOptions || '<option value="">لا توجد رتب متاحة</option>'}</select></div>
                <div class="col-md-3"><label class="form-label">نوع الحد</label><select id="rank_limit_target" class="form-select">${rankLimitTargets}</select></div>
                <div class="col-md-2"><label class="form-label">الحد</label><input id="rank_limit_value" type="number" min="0" class="form-control"></div>
                <div class="col-md-2 text-end"><button class="btn btn-outline-primary" onclick="applyRankLimit()">تطبيق</button></div>
              </div>
              <div class="mt-2"><label class="form-label">استثناء ضباط</label><select id="rank_limit_excluded" multiple class="form-select" style="min-height:88px" onchange="updateRankLimitExclusions()">${rankLimitExcludedOptions}</select></div>
              <div class="small-muted mt-1">لن يتم تطبيق الحد على الرتب المخفية أو الضباط المستثنين.</div>
             </div>`+
             `<p class="small-muted">اترك الخانة فارغة لعدم وجود حد. خانة الإجمالي تضبط إجمالي الخدمات لكل الضابط في الشهر.</p>`+
             `<div class="table-responsive"><table class="table table-bordered table-sm"><thead class="table-dark"><tr><th>الرتبة</th><th>الاسم</th>${dutyHeaders}<th>الإجمالي الشهري</th></tr></thead><tbody>${rows || '<tr><td colspan="${duties.length+2}" class="text-center text-muted">لا يوجد ضباط بعد تطبيق عوامل الإخفاء</td></tr>'}</tbody></table></div></div></div>`;
    }

    function toggleLimitRankFilter(rank, hidden){
      const set = new Set(officerLimitFilters.hiddenRanks || []);
      if(hidden) set.add(rank); else set.delete(rank);
      officerLimitFilters.hiddenRanks = Array.from(set);
      saveAll();
      renderTab('limits');
    }
    function setLimitDeptVisibility(key, hidden){
      if(key==='internal') officerLimitFilters.hideInternal = hidden;
      if(key==='external') officerLimitFilters.hideExternal = hidden;
      saveAll();
      renderTab('limits');
    }
    function updateRankLimitExclusions(){
      const selected = Array.from(document.getElementById('rank_limit_excluded')?.selectedOptions || []).map(opt=>+opt.value);
      officerLimitFilters.rankLimitExcluded = selected;
      saveAll();
    }
    function officerHasExceptionRule(officerId){
      return (exceptions || []).some(rule => rule.targetType === 'officer' && rule.targetId === officerId);
    }
    function applyRankLimit(){
      const rank = document.getElementById('rank_limit_rank')?.value || '';
      const target = document.getElementById('rank_limit_target')?.value || 'total';
      const rawVal = document.getElementById('rank_limit_value')?.value;
      if(!rank) return showToast('اختر رتبة لتطبيق الحد', 'warning');
      if(officerLimitFilters.hiddenRanks.includes(rank)) return showToast('الرتبة مخفية ولن يتم تطبيق الحد عليها', 'warning');
      const normalized = normalizeLimit(rawVal);
      const excluded = new Set((officerLimitFilters.rankLimitExcluded || []).map(Number));
      let updated = 0;
      officers.forEach(o=>{
        if(o.rank !== rank) return;
        if(officerLimitFilters.hiddenRanks.includes(o.rank)) return;
        if(excluded.has(o.id)) return;
        if(officerHasExceptionRule(o.id)) return;
        ensureOfficerLimit(o.id);
        if(target === 'total'){
          officerLimits[o.id].total = normalized;
        } else {
          const dutyId = +target;
          officerLimits[o.id].duties[dutyId] = normalized;
        }
        updated++;
      });
      saveAll();
      renderTab('limits');
      if(updated){
        showToast(`تم تطبيق الحد على ${updated} ضابط`, 'success');
      } else {
        showToast('لا يوجد ضباط مطابقون للتطبيق', 'warning');
      }
    }

    /* ========= Duties ========= */
    function renderDutiesTab(){
      if(SETTINGS.authEnabled && !CURRENT_USER) return signInInlineFormMarkup() + '<div class="alert alert-warning">سجّل الدخول للوصول لأنواع الخدمات.</div>';
      return `<div class="card"><div class="card-header d-flex justify-content-between"><div>أنواع الخدمات</div><div><button class="btn btn-success btn-sm" onclick="addDuty()">إضافة نوع</button></div></div>
        <div class="card-body">
          <div id="dutiesEditor">${duties.map(d=>renderDutyEditor(d)).join('')}</div>
        </div></div>`;
    }
    function renderDutyEditor(d){
      const allowedRanksHtml = ranks.map(r=>`<div class="form-check"><input class="form-check-input" type="checkbox" id="dr_${d.id}_${r}" ${d.allowedRanks && d.allowedRanks.includes(r)?'checked':''} onchange="toggleDutyRank(${d.id},'${r}',this.checked)"><label class="form-check-label" for="dr_${d.id}_${r}">${r}</label></div>`).join('');
      return `<div class="border rounded p-3 mb-3" id="duty_editor_${d.id}">
        <div class="row g-2 align-items-center">
          <div class="col-md-3"><label class="form-label">الاسم</label><input class="form-control" value="${escapeHtml(d.name)}" onchange="updateDuty(${d.id},'name',this.value)"></div>
          <div class="col-md-3"><label class="form-label">وضعية الطباعة (label)</label><input class="form-control" value="${escapeHtml(d.printLabel||d.name)}" onchange="updateDuty(${d.id},'printLabel',this.value)"></div>
          <div class="col-md-2"><label class="form-label">اللون</label><select class="form-select" onchange="updateDuty(${d.id},'color',this.value)">${['danger','success','warning','dark','primary','info'].map(c=>`<option value="${c}" ${c===d.color?'selected':''}>${c}</option>`).join('')}</select></div>
          <div class="col-md-1"><label class="form-label">ساعات</label><input type="number" class="form-control" value="${d.duration}" onchange="updateDuty(${d.id},'duration',+this.value)"></div>
          <div class="col-md-1"><label class="form-label">راحة (س)</label><input type="number" class="form-control" value="${d.restHours}" onchange="updateDuty(${d.id},'restHours',+this.value)"></div>
          <div class="col-md-1 text-end"><label class="form-label d-block">&nbsp;</label><button class="btn btn-danger" onclick="if(confirm('حذف؟')) deleteDuty(${d.id})">حذف</button></div>
        </div>

       <div class="mt-3"><label class="form-label fw-bold">الرتب المسموحة</label><div class="compact-checks">${allowedRanksHtml}</div></div>
      </div>`;
    }
    function addDuty(){ duties.push({id:Date.now(),name:"نوع مناوبة",printLabel:"نوع مناوبة",color:"primary",duration:8,restHours:24,allowedRanks:[],signingJobTitleIds:[]}); saveAll(); renderTab('duties'); }
    function updateDuty(id,key,val){ const d = duties.find(x=>x.id===id); if(!d) return; d[key] = val; saveAll(); renderTab('duties'); }
    function toggleDutyRank(dutyId,rank,checked){ const d = duties.find(x=>x.id===dutyId); if(!d) return; d.allowedRanks = d.allowedRanks || []; if(checked){ if(!d.allowedRanks.includes(rank)) d.allowedRanks.push(rank); } else d.allowedRanks = d.allowedRanks.filter(r=>r!==rank); saveAll(); renderTab('duties'); }
    function deleteDuty(id){ duties = duties.filter(d=>d.id!==id); saveAll(); renderTab('duties'); }
    function toggleDutySigning(){ /* توقيع لكل نوع غير مدعوم بعد نقل إدارة التوقيع */ }

    /* ========= Exceptions ========= */
    function renderExceptionsTab(){
      if(SETTINGS.authEnabled && (!CURRENT_USER || currentRole()==='editor')) return signInInlineFormMarkup() + '<div class="alert alert-warning">ليست لديك صلاحية تعديل القواعد.</div>';
      return `<div class="card"><div class="card-header">القواعد والاستثناءات</div><div class="card-body">
        <div id="exceptionsFormArea">${renderExceptionForm()}</div>
        <h5 class="mt-3">القواعد الحالية</h5>
        <div id="exceptionsList">${exceptions.length? exceptions.map((e,i)=>renderExceptionRow(e,i)).join('') : '<p class="text-muted py-2">لا توجد قواعد</p>'}</div>
      </div></div>`;
    }
    function renderExceptionForm(){
      return `<div class="card mb-3"><div class="card-body">
        <div class="row g-2">
          <div class="col-md-3"><label class="form-label">النوع</label><select id="ex_type" class="form-select">
            <option value="force_weekly_only">حجز يوم ثابت (حصري)</option>
            <option value="force_weekday">فرض يوم/تاريخ محدد</option>
            <option value="force_duty">فرض خدمة معينة</option>
            <option value="deny_duty">منع خدمة معينة</option>
            <option value="remove_all">إزالة من كل الجداول</option>
            <option value="vacation">إجازة</option>
            <option value="weekday">منع يوم أسبوعي</option>
          </select></div>
          <div class="col-md-3"><label class="form-label">نوع المناوبة</label><select id="ex_duty" class="form-select"><option value="">كل الخدمات</option>${duties.map(d=>`<option value="${d.id}">${escapeHtml(d.name)}</option>`).join('')}</select></div>
          <div class="col-md-3"><label class="form-label">المستهدف</label><select id="ex_targetType" class="form-select"><option value="officer">ضابط</option><option value="dept">قسم</option><option value="rank">رتبة</option></select></div>
          <div class="col-md-3" id="ex_targetContainer"></div>
          <div class="col-md-3" id="ex_from_container" style="display:none;"><label class="form-label">من</label><input id="ex_from" type="date" class="form-control"></div>
          <div class="col-md-3" id="ex_to_container" style="display:none;"><label class="form-label">إلى</label><input id="ex_to" type="date" class="form-control"></div>
          <div class="col-md-3" id="ex_weekday_container" style="display:none;">
            <label class="form-label">أيام أسبوعية (يمكن اختيار أكثر من يوم)</label>
            <select id="ex_weekdays" class="form-select" multiple>${weekdays.map((d,i)=>`<option value="${i}">${d}</option>`).join('')}</select>
            <label class="form-label mt-2">طريقة التعامل مع تعدد الأيام</label>
            <select id="ex_weekday_mode" class="form-select">
              <option value="all">تطبيق جميع الأيام المختارة</option>
              <option value="alternate">تناوب أسبوعي (يوم واحد بالأسبوع)</option>
            </select>
            <small class="text-muted">عند اختيار أكثر من يوم يمكن تفعيل التناوب الأسبوعي.</small>
          </div>
          <div class="col-md-3" id="ex_day_container" style="display:none;"><label class="form-label">تاريخ محدد (اختياري)</label><input id="ex_day" type="number" min="1" max="31" class="form-control" placeholder="مثال: 15"></div>
          <div class="col-12"><label class="form-label">استثناء ضباط</label><select id="ex_excludedOfficers" multiple class="form-select" style="min-height:100px">${officers.map(o=>`<option value="${o.id}">${escapeHtml(o.rank)} ${escapeHtml(o.name)}</option>`).join('')}</select></div>
          <div class="col-12 text-end"><button class="btn btn-secondary" onclick="clearExceptionForm()">مسح</button> <button class="btn btn-primary" onclick="saveExceptionFromForm()">حفظ</button></div>
        </div>
      </div></div>`;
    }
    function renderExceptionRow(e,i){
      const target = e.targetType==='rank'? e.targetId : (e.targetType==='dept'? departments.find(d=>d.id===e.targetId)?.name : officers.find(o=>o.id===e.targetId)?.name);
      const dutyLabel = e.dutyId ? (duties.find(d=>d.id===e.dutyId)?.name || '') : 'كل الخدمات';
      const typeLabel = {
        force_weekday: 'فرض يوم محدد',
        force_weekly_only: 'حجز يوم حصري',
        force_duty: 'فرض خدمة',
        deny_duty: 'منع خدمة',
        remove_all: 'إزالة من كل الجداول',
        vacation: 'إجازة',
        weekday: 'منع أسبوعي',
        block: 'منع خدمة'
      }[e.type] || e.type;
      let meta = `${typeLabel} — ${dutyLabel} — ${target||''}`;
      if(e.type==='weekday' || e.type==='force_weekday' || e.type==='force_weekly_only'){
        const weekLabel = Array.isArray(e.weekdays) && e.weekdays.length
          ? e.weekdays.map(w=>weekdays[+w]).join('، ')
          : weekdays[e.weekday];
        const modeLabel = e.weekdayMode === 'alternate' && Array.isArray(e.weekdays) && e.weekdays.length > 1 ? ' (تناوب أسبوعي)' : '';
        meta += ` | ${weekLabel}${modeLabel}`;
      }
      if(e.dayOfMonth) meta += ` | يوم ${e.dayOfMonth}`;
      if(e.type==='vacation') meta += ` | ${e.fromDate} → ${e.toDate}`;
      if(e.excludedOfficers?.length) meta += `<br/><small>مستثنون: ${e.excludedOfficers.map(id=>officers.find(o=>o.id===id)?.name||id).join(', ')}</small>`;
      return `<div class="p-2 mb-2 border rounded d-flex justify-content-between align-items-center">${meta}
        <div><button class="btn btn-sm btn-outline-primary me-1" onclick="populateExceptionForm(${i})">تعديل</button><button class="btn btn-sm btn-outline-danger" onclick="deleteException(${i})">حذف</button></div></div>`;
    }
    function deleteException(i){ exceptions.splice(i,1); saveAll(); renderTab('exceptions'); }
    function clearExceptionForm(){
      document.getElementById('ex_type').value='force_weekly_only';
      document.getElementById('ex_duty').value='';
      document.getElementById('ex_targetType').value='officer';
      populateExceptionTarget('officer',null);
      document.getElementById('ex_from').value=''; document.getElementById('ex_to').value='';
      Array.from(document.getElementById('ex_weekdays')?.options||[]).forEach(o=>o.selected=false);
      if(document.getElementById('ex_weekday_mode')) document.getElementById('ex_weekday_mode').value='all';
      document.getElementById('ex_day').value='';
      Array.from(document.getElementById('ex_excludedOfficers').options || []).forEach(o=>o.selected=false);
      const edit = document.getElementById('ex_editIndex'); if(edit) edit.remove();
      toggleExceptionFields();
    }
    function populateExceptionTarget(type,value){
      const c = document.getElementById('ex_targetContainer');
      if(!c) return;
      if(type==='officer') c.innerHTML = `<label class="form-label">اختر ضابط</label><select id="ex_targetId" class="form-select"><option value="">اختر</option>${officers.map(o=>`<option value="${o.id}" ${o.id==value?'selected':''}>${escapeHtml(o.rank)} ${escapeHtml(o.name)}</option>`).join('')}</select>`;
      else if(type==='dept') c.innerHTML = `<label class="form-label">اختر قسم</label><select id="ex_targetId" class="form-select"><option value="">اختر</option>${departments.map(d=>`<option value="${d.id}" ${d.id==value?'selected':''}>${escapeHtml(d.name)}</option>`).join('')}</select>`;
      else if(type==='rank') c.innerHTML = `<label class="form-label">اختر رتبة</label><select id="ex_targetId" class="form-select">${ranks.map(r=>`<option ${r==value?'selected':''}>${r}</option>`).join('')}</select>`;
      else c.innerHTML='';
    }
    function toggleExceptionFields(){
      const t = document.getElementById('ex_type').value;
      const showVacation = t==='vacation';
      const showWeekday = t==='weekday' || t==='force_weekday' || t==='force_weekly_only';
      const showDay = t==='force_weekday' || t==='force_weekly_only';
      document.getElementById('ex_from_container').style.display = showVacation?'block':'none';
      document.getElementById('ex_to_container').style.display = showVacation?'block':'none';
      document.getElementById('ex_weekday_container').style.display = showWeekday?'block':'none';
      document.getElementById('ex_day_container').style.display = showDay?'block':'none';
      const dutySelect = document.getElementById('ex_duty');
      if(dutySelect){ dutySelect.disabled = (t==='remove_all'); if(t==='remove_all') dutySelect.value=''; }
      updateWeekdayModeState();
    }
    function updateWeekdayModeState(){
      const weekdaysSelect = document.getElementById('ex_weekdays');
      const modeSelect = document.getElementById('ex_weekday_mode');
      if(!weekdaysSelect || !modeSelect) return;
      const selected = Array.from(weekdaysSelect.selectedOptions || []).map(o=>o.value);
      if(selected.length <= 1){
        modeSelect.value = 'all';
        modeSelect.disabled = true;
      } else {
        modeSelect.disabled = false;
      }
    }
    document.addEventListener('change', e=>{
      if(e.target && e.target.id==='ex_targetType') populateExceptionTarget(e.target.value,null);
      if(e.target && e.target.id==='ex_type') toggleExceptionFields();
      if(e.target && e.target.id==='ex_weekdays') updateWeekdayModeState();
    });
    function saveExceptionFromForm(){
      const type = document.getElementById('ex_type').value;
      const dutyId = document.getElementById('ex_duty').value ? +document.getElementById('ex_duty').value : null;
      const targetType = document.getElementById('ex_targetType').value;
      const targetEl = document.getElementById('ex_targetId');
      const targetId = targetEl ? (targetType==='rank'? targetEl.value : (targetEl.value? +targetEl.value : null)) : null;
      const excluded = Array.from(document.getElementById('ex_excludedOfficers').selectedOptions || []).map(o=>+o.value);
      const rule = {type,dutyId,targetType,targetId,excludedOfficers:excluded};
      if(targetId==null) return showToast('حدد المستهدف أولاً','danger');
      if(type==='vacation'){ rule.fromDate = document.getElementById('ex_from').value; rule.toDate = document.getElementById('ex_to').value; if(!rule.fromDate||!rule.toDate) return showToast('اختر تواريخ صحيحة','danger'); }
      if(type==='weekday' || type==='force_weekday' || type==='force_weekly_only'){
        const weekdayVals = Array.from(document.getElementById('ex_weekdays')?.selectedOptions||[]).map(o=>+o.value);
        rule.weekdays = weekdayVals;
        rule.weekday = weekdayVals.length ? weekdayVals[0] : null;
        const mode = document.getElementById('ex_weekday_mode')?.value || 'all';
        rule.weekdayMode = weekdayVals.length > 1 ? mode : 'all';
        rule.dayOfMonth = document.getElementById('ex_day').value ? +document.getElementById('ex_day').value : null;
      }
      if(type==='force_duty' || type==='deny_duty'){ if(!dutyId) return showToast('اختر نوع المناوبة المستهدف','danger'); }
      if(type==='force_weekly_only' && targetType==='officer' && (!rule.weekdays || !rule.weekdays.length)) return showToast('حدد يوم الأسبوع المطلوب للحجز','danger');
      if(type==='remove_all'){ rule.dutyId = null; }
      const editIndex = document.getElementById('ex_editIndex') ? +document.getElementById('ex_editIndex').value : null;
      const wasEdit = editIndex!=null && !isNaN(editIndex);
      if(wasEdit) exceptions[editIndex] = rule; else exceptions.push(rule);
      saveAll();
      showToast(wasEdit ? 'تم حفظ التعديل اليدوي على القاعدة' : 'تم حفظ القاعدة الجديدة', wasEdit ? 'warning' : 'success');
      renderTab('exceptions');
    }
    function populateExceptionForm(i){
      const e = exceptions[i]; if(!e) return;
      document.getElementById('ex_type').value = e.type || 'block';
      document.getElementById('ex_duty').value = e.dutyId || '';
      document.getElementById('ex_targetType').value = e.targetType || 'officer';
      populateExceptionTarget(e.targetType,e.targetId);
      document.getElementById('ex_from').value = e.fromDate || '';
      document.getElementById('ex_to').value = e.toDate || '';
      const weekdaysSelect = document.getElementById('ex_weekdays');
      if(weekdaysSelect){
        const selected = Array.isArray(e.weekdays) && e.weekdays.length ? e.weekdays.map(Number) : [e.weekday ?? 0];
        Array.from(weekdaysSelect.options).forEach(opt=> opt.selected = selected.includes(+opt.value));
      }
      if(document.getElementById('ex_weekday_mode')){
        document.getElementById('ex_weekday_mode').value = e.weekdayMode || 'all';
      }
      document.getElementById('ex_day').value = e.dayOfMonth || '';
      Array.from(document.getElementById('ex_excludedOfficers').options || []).forEach(opt=> opt.selected = Array.isArray(e.excludedOfficers) && e.excludedOfficers.includes(+opt.value));
      if(!document.getElementById('ex_editIndex')){ const inp=document.createElement('input'); inp.type='hidden'; inp.id='ex_editIndex'; document.getElementById('exceptionsFormArea').appendChild(inp); }
      document.getElementById('ex_editIndex').value = i;
      toggleExceptionFields();
      window.scrollTo({top:0,behavior:'smooth'});
    }

    /* ========= Relations Diagram ========= */
    function renderRelationsTab(){
      if(SETTINGS.authEnabled && !CURRENT_USER) return signInInlineFormMarkup() + '<div class="alert alert-warning">سجّل الدخول لعرض المخطط.</div>';
      const deptCards = departments.map(d=>{
        const parent = departments.find(p=>p.id===d.parentId);
        const head = d.headId ? officers.find(o=>o.id===d.headId) : null;
        const upperOfficer = d.upperOfficerId ? officers.find(o=>o.id===d.upperOfficerId) : null;
        const deptOfficers = officers.filter(o=>o.deptId===d.id);
        const officersList = deptOfficers.length
          ? `<ul class="relation-list">${deptOfficers.map(o=>`<li><span>${escapeHtml(o.rank)} ${escapeHtml(o.name)}</span><span class="relation-tag">${escapeHtml(o.status || 'active')}</span></li>`).join('')}</ul>`
          : '<div class="text-muted">لا يوجد ضباط بالقسم.</div>';
        return `<div class="relation-card">
          <h5>${escapeHtml(d.name)}</h5>
          <div class="relation-item"><span>يتبع</span><span class="relation-tag">${parent ? escapeHtml(parent.name) : '—'}</span></div>
          <div class="relation-item"><span>رئيس القسم</span><span class="relation-tag">${head ? escapeHtml(head.name) : 'غير محدد'}</span></div>
          ${d.upperTitle || upperOfficer ? `<div class="relation-item"><span>إدارة عليا</span><span class="relation-tag">${escapeHtml(d.upperTitle || '')}${upperOfficer ? ` — ${escapeHtml(upperOfficer.name)}` : ''}</span></div>` : ''}
          <div class="relation-item"><span>الضباط</span><span class="relation-tag">${deptOfficers.length}</span></div>
          <div class="mt-2">${officersList}</div>
        </div>`;
      }).join('');
      const titleCards = jobTitles.map(t=>{
        const parent = jobTitles.find(p=>p.id===t.parentId);
        const count = officers.filter(o=>o.jobTitleId===t.id).length;
        return `<div class="relation-card">
          <h5>${escapeHtml(t.name)}</h5>
          <div class="relation-item"><span>يتبع</span><span class="relation-tag">${parent ? escapeHtml(parent.name) : '—'}</span></div>
          <div class="relation-item"><span>ضباط</span><span class="relation-tag">${count}</span></div>
          <div class="relation-item"><span>قيادي</span><span class="relation-tag">${t.isUpper ? 'نعم' : 'لا'}</span></div>
        </div>`;
      }).join('');
     return `<div class="card">
        <div class="card-header d-flex justify-content-between align-items-center">
          <div>Organizational structure — الأقسام والمسميات والضباط</div>
          <div class="small text-muted">عرض مخطط تنظيمي للهيكل الإداري</div>
        </div>
        <div class="card-body">
          <div class="relations-board">
            <div class="relations-column">
              <h5 class="mb-2">الأقسام</h5>
              ${deptCards || '<div class="text-muted">لا توجد أقسام مسجلة.</div>'}
            </div>
            <div class="relations-column">
              <h5 class="mb-2">المسميات الوظيفية</h5>
              ${titleCards || '<div class="text-muted">لا توجد مسميات مسجلة.</div>'}
            </div>
          </div>
        </div>
      </div>`;
    }

    /* ========= Report ========= */
    function renderReportTab(){
      if(SETTINGS.authEnabled && !CURRENT_USER) return signInInlineFormMarkup() + '<div class="alert alert-warning">سجّل الدخول لرؤية التقارير.</div>';
      const month = document.getElementById('rosterMonth')?.value || sessionStorage.getItem('activeRosterMonth') || new Date().toISOString().slice(0,7);
      const title = new Date(month+'-01').toLocaleDateString('ar-EG',{month:'long',year:'numeric'});
     let savedSelection = [];
      try { savedSelection = JSON.parse(sessionStorage.getItem('selectedReportDutyIds')) || []; } catch(_) { savedSelection = []; }
      const dutyOptions = duties.map(d=>`<div class="form-check form-check-inline"><input class="form-check-input rpt-duty-chk" type="checkbox" id="rpd_${d.id}" value="${d.id}" ${savedSelection.includes(d.id)?'checked':''} onchange="cacheReportSelection()"><label class="form-check-label" for="rpd_${d.id}">${escapeHtml(d.printLabel||d.name)}</label></div>`).join(' ');
        return `<div class="card"><div class="card-header d-flex justify-content-between"><div>التقرير — ${title}</div>
        <div>
          <button class="btn btn-outline-primary btn-sm me-2" onclick="printAllDuties()">طباعة كل الأنواع</button>
          <button class="btn btn-outline-secondary btn-sm" onclick="exportDutiesByFormat('all')">تصدير الكل</button>
        </div></div>
        <div class="card-body">
          <div class="row mb-3">
            <div class="col-md-4"><label class="form-label">شهر</label><input id="rpt_month" type="month" class="form-control" value="${month}" onchange="document.getElementById('rosterMonth').value=this.value; onRosterMonthChange();"></div>
            <div class="col-md-8"><label class="form-label">اختر أنواع للطباعة/التصدير</label><div class="compact-checks">${dutyOptions}</div></div>
          </div>
          <div class="card mb-3">
            <div class="card-header">إعدادات الخط والتخطيط</div>
            <div class="card-body">
              <div class="row g-2">
                <div class="col-md-4"><label class="form-label">خط التقرير</label><select id="report_font_family" class="form-select">${fontOptions.map(f=>`<option value="${f.value}" ${SETTINGS.reportFontFamily===f.value?'selected':''}>${f.label}</option>`).join('')}</select></div>
                <div class="col-md-3"><label class="form-label">حجم الخط</label><input id="report_font_size" type="number" class="form-control" value="${SETTINGS.printFontSize}" min="8" max="18"></div>
                <div class="col-md-5"><label class="form-label">رفع خط مخصص (.ttf/.otf)</label><input id="report_custom_font" type="file" class="form-control" accept=".ttf,.otf"></div>
              </div>
              <div class="mt-3 text-end"><button class="btn btn-outline-primary" onclick="saveReportSettingsFromTab()">حفظ إعدادات التقرير</button></div>
            </div>
          </div>
          <div class="row g-2 mb-3 align-items-end">
            <div class="col-md-4"><label class="form-label">امتداد التصدير</label><select id="rpt_export_format" class="form-select"><option value="pdf">PDF</option><option value="png">PNG</option><option value="jpeg">JPEG</option><option value="html">HTML</option><option value="docx">DOCX</option></select></div>
            <div class="col-md-8 text-end">
              <button class="btn btn-success me-2" onclick="previewSelectedDuties()">معاينة</button>
              <button class="btn btn-outline-primary me-2" onclick="exportDutiesByFormat('selected')">تصدير المحدد</button>
              <button class="btn btn-outline-secondary" onclick="printSelectedDuties()">طباعة المحدد</button>
            </div>
          </div>
          <div id="reportPreview" class="print-preview"></div>
        </div></div>`;
    }
    function getSelectedDutyIdsInReport(){
      return Array.from(document.querySelectorAll('.rpt-duty-chk:checked')).map(i=>+i.value);
    }
    function cacheReportSelection(){
      const ids = getSelectedDutyIdsInReport();
      sessionStorage.setItem('selectedReportDutyIds', JSON.stringify(ids));
    }
    function previewSelectedDuties(){
      const ids = getSelectedDutyIdsInReport();
      const month = document.getElementById('rpt_month')?.value || sessionStorage.getItem('activeRosterMonth') || new Date().toISOString().slice(0,7);
      if(!ids.length) return showToast('اختر نوعاً واحداً أو أكثر للمعاينة','danger');
      const preview = document.getElementById('reportPreview');
      preview.innerHTML = ids.map(id=> buildDutyInnerHtml(id, month)).join('<hr/>');
      preview.scrollIntoView({behavior:'smooth'});
    }
    function saveReportSettingsFromTab(){
      SETTINGS.reportFontFamily = document.getElementById('report_font_family')?.value || SETTINGS.reportFontFamily;
      SETTINGS.printFontSize = +document.getElementById('report_font_size')?.value || SETTINGS.printFontSize;
      const fontFile = document.getElementById('report_custom_font')?.files?.[0];
      if(fontFile){
        const fr = new FileReader();
        fr.onload = ()=>{
          SETTINGS.customFontData = fr.result;
          SETTINGS.reportFontFamily = 'AppCustomFont';
          applyCustomFont();
          saveAll();
          showToast('تم تحديث إعدادات الخط للتقرير','success');
          renderTab('report');
        };
        fr.readAsDataURL(fontFile);
        return;
      }
      saveAll();
      showToast('تم حفظ إعدادات التقرير','success');
      renderTab('report');
    }
    function downloadTextFile(name, content, type='text/plain'){
      const blob=new Blob([content],{type:`${type};charset=utf-8`});
      const a=document.createElement('a');
      a.href=URL.createObjectURL(blob);
      a.setAttribute('download',name);
      document.body.appendChild(a);
      a.click();
      a.remove();
    }
    function buildReportHtmlDocument(ids, month){
      const fontFamily = getReportFontFamily();
      const fontSize = getReportPrintFontSize();
      const bodyContent = ids.map(id=> buildDutyInnerHtml(id, month)).join('<div style="page-break-after:always;"></div>');
      return `<!doctype html><html lang="ar" dir="rtl"><head><meta charset="utf-8"><title>Roster ${month}</title>
        <style>
          ${buildReportStyles({fontFamily, fontSize, includeBackground: true})}
          .report-preview-paper{background:#fff;border:1px solid #e5e7eb;border-radius:10px;padding:16px;box-shadow:0 8px 20px rgba(15,23,42,.08)}
          .report-header{display:flex;justify-content:space-between;align-items:center;gap:12px;border-bottom:2px solid #1e40af;padding-bottom:8px;margin-bottom:10px}
          .report-meta{display:flex;justify-content:space-between;align-items:center;gap:12px;font-size:12px;color:#334155;margin-bottom:8px}
          .report-title{font-size:18px;font-weight:800;text-align:center;color:#1e40af}
        </style>
      </head><body>${bodyContent}</body></html>`;
    }
    function buildReportDocxDocument(ids, month){
      const fontFamily = getReportFontFamily();
      const fontSize = getReportPrintFontSize();
      const bodyContent = ids.map(id=> buildDutyInnerHtml(id, month)).join('<div style="page-break-after:always;"></div>');
      return `<!doctype html><html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns="http://www.w3.org/TR/REC-html40" lang="ar" dir="rtl"><head><meta charset="utf-8">
        <title>Roster ${month}</title>
        <!--[if gte mso 9]><xml><w:WordDocument><w:View>Print</w:View><w:Zoom>90</w:Zoom></w:WordDocument></xml><![endif]-->
        <style>
          ${buildReportStyles({fontFamily, fontSize, includeBackground: false})}
          .report-preview-paper{background:#fff;border:1px solid #e5e7eb;border-radius:10px;padding:16px}
          .report-header{display:flex;justify-content:space-between;align-items:center;gap:12px;border-bottom:2px solid #1e40af;padding-bottom:8px;margin-bottom:10px}
          .report-meta{display:flex;justify-content:space-between;align-items:center;gap:12px;font-size:12px;color:#334155;margin-bottom:8px}
          .report-title{font-size:18px;font-weight:800;text-align:center;color:#1e40af}
        </style>
      </head><body>${bodyContent}</body></html>`;
    }
    async function exportDutiesByFormat(scope='selected'){
      const month = document.getElementById('rpt_month')?.value || sessionStorage.getItem('activeRosterMonth') || new Date().toISOString().slice(0,7);
      const ids = scope === 'all' ? duties.map(d=>d.id) : getSelectedDutyIdsInReport();
      if(!ids.length) return showToast('اختر نوعاً واحداً أو أكثر للتصدير','danger');
      const format = document.getElementById('rpt_export_format')?.value || 'pdf';
      if(format === 'pdf'){
        if(scope === 'all') return saveAllDutiesPDF();
        return saveSelectedDutiesPDF();
      }
      if(format === 'html'){
        const html = buildReportHtmlDocument(ids, month);
        downloadTextFile(`Roster_${scope}_${month}.html`, html, 'text/html');
        return showToast('تم تصدير ملف HTML','success');
      }
      if(format === 'docx'){
        const doc = buildReportDocxDocument(ids, month);
        downloadTextFile(`Roster_${scope}_${month}.docx`, doc, 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
        return showToast('تم تصدير ملف DOCX','success');
      }
      if(format === 'docx'){
        const doc = buildReportDocxDocument(ids, month);
        downloadTextFile(`Roster_${scope}_${month}.docx`, doc, 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
        return showToast('تم تصدير ملف DOCX','success');
      }
      if(format === 'png' || format === 'jpeg'){
        const container = document.createElement('div');
        container.style.width = '210mm';
        container.style.padding = '6mm';
        container.style.direction = 'rtl';
        container.innerHTML = ids.map(id=> buildDutyInnerHtml(id, month)).join('<div style="page-break-after:always;"></div>');
        document.body.appendChild(container);
        try{
          await ensureHtml2CanvasReady();
          const canvas = await html2canvas(container, {scale:2, useCORS:true});
          const mime = format === 'png' ? 'image/png' : 'image/jpeg';
          const dataUrl = canvas.toDataURL(mime, 0.95);
          const link = document.createElement('a');
          link.href = dataUrl;
          link.download = `Roster_${scope}_${month}.${format}`;
          link.click();
          showToast(`تم تصدير صورة ${format.toUpperCase()}`,'success');
        }catch(err){
          console.error(err);
          showToast('فشل إنشاء الصورة','danger');
        }finally{
          container.remove();
        }
        return;
      }
      showToast('صيغة التصدير غير مدعومة حالياً','warning');
  }

    /* ========= Settings ========= */
    function getSettingsSubTab(){
      return sessionStorage.getItem('settingsSubTab') || 'general';
    }
    function setSettingsSubTab(id){
      sessionStorage.setItem('settingsSubTab', id);
      renderTab('settings');
    }
    function exportSettingsBackup(){
      const payload = { settings: SETTINGS, meta:{ type:'settings', exportedAt: new Date().toISOString() } };
      downloadJSONFile(`settings_backup_${new Date().toISOString().slice(0,10)}.json`, JSON.stringify(payload, null, 2));
      showToast('تم تنزيل نسخة الإعدادات','success');
    }
    function importSettingsBackup(file){
      if(!file) return;
      const reader = new FileReader();
      reader.onload = ()=>{
        try{
          const data = JSON.parse(reader.result);
          if(!data.settings) throw new Error('Invalid settings backup');
          SETTINGS = Object.assign({}, defaultSettings, data.settings || {});
          saveAll();
          loadAll();
          applyTheme();
          applyCustomFont();
          applyFontSettings();
          updateNavLogo();
          renderTab('settings');
          showToast('تمت استعادة إعدادات النظام','success');
        }catch(err){
          console.error(err);
          showToast('ملف الإعدادات غير صالح','danger');
        }
      };
      reader.readAsText(file);
    }
    function renderSettingsTab(){
      if(SETTINGS.authEnabled && !CURRENT_USER) return signInInlineFormMarkup() + '<div class="alert alert-warning">سجّل الدخول لإدارة الإعدادات.</div>';
      const active = getSettingsSubTab();
      const tabs = [
        {id:'general', label:'عام'},
        {id:'appearance', label:'المظهر والخطوط'},
        {id:'reports', label:'التقارير والطباعة'},
        {id:'signatures', label:'التوقيعات'},
        {id:'backup', label:'نسخ/استعادة'}
      ];
      const tabButtons = `<div class="sub-tabs">${tabs.map(t=>`<button class="sub-tab-btn ${active===t.id?'active':''}" onclick="setSettingsSubTab('${t.id}')">${t.label}</button>`).join('')}</div>`;
      const sections = {
        general: `
          <div class="row g-2">
            <div class="col-md-6"><label class="form-label">اسم التطبيق</label><input id="set_appName" class="form-control" value="${escapeHtml(SETTINGS.appName)}"></div>
            <div class="col-md-6"><label class="form-label">عنوان ثانوي</label><input id="set_appSubtitle" class="form-control" value="${escapeHtml(SETTINGS.appSubtitle)}"></div>
            <div class="col-md-6"><label class="form-label">تخطيط شاشة الدخول</label><select id="set_loginLayout" class="form-select"><option value="stacked" ${SETTINGS.loginLayout==='stacked'?'selected':''}>مكدس (شعار فوق البطاقة)</option><option value="inline" ${SETTINGS.loginLayout==='inline'?'selected':''}>عرض مضغوط</option></select></div>
            <div class="col-md-6 form-check align-items-center d-flex"><input id="set_auth" class="form-check-input" type="checkbox" ${SETTINGS.authEnabled?'checked':''}><label class="form-check-label ms-2">تمكين نظام المستخدمين</label></div>
           <div class="col-md-6"><label class="form-label">خادم المزامنة (API)</label><input id="set_apiBaseUrl" class="form-control" placeholder="https://example.com" value="${escapeHtml(SETTINGS.apiBaseUrl || '')}"></div>
            <div class="col-md-6 form-check align-items-center d-flex"><input id="set_apiSync" class="form-check-input" type="checkbox" ${SETTINGS.apiSyncEnabled?'checked':''}><label class="form-check-label ms-2">تمكين المزامنة عبر الخادم</label></div>
            <div class="col-12"><div class="form-text text-muted">استخدم نفس النطاق أو اترك الحقل فارغًا لاستخدام نطاق الصفحة الحالية. عند التعطيل سيتم استخدام التخزين المحلي فقط.</div></div>
           </div>
        `,
        appearance: `
          <div class="row g-2">
            <div class="col-md-4"><label class="form-label">شعار (تحميل)</label><input id="set_logo" type="file" class="form-control" accept="image/*"></div>
            <div class="col-md-4"><label class="form-label">السمة</label><select id="set_theme" class="form-select">${Object.keys(THEMES).map(k=>`<option value="${k}" ${SETTINGS.themeKey===k?'selected':''}>${THEMES[k].name}</option>`).join('')}</select></div>
            <div class="col-md-4"><label class="form-label">خط التطبيق</label><select id="set_appFont" class="form-select">${fontOptions.map(f=>`<option value="${f.value}" ${SETTINGS.appFontFamily===f.value?'selected':''}>${f.label}</option>`).join('')}</select></div>
            <div class="col-md-4"><label class="form-label">خط الحقول النصية</label><select id="set_formFont" class="form-select">${fontOptions.map(f=>`<option value="${f.value}" ${SETTINGS.formFontFamily===f.value?'selected':''}>${f.label}</option>`).join('')}</select></div>
            <div class="col-md-4"><label class="form-label">خط التبويبات</label><select id="set_tabFont" class="form-select">${fontOptions.map(f=>`<option value="${f.value}" ${SETTINGS.tabFontFamily===f.value?'selected':''}>${f.label}</option>`).join('')}</select></div>
          </div>
        `,
        reports: `
          <div class="row g-2">
            <div class="col-12"><label class="form-label">رأس التقرير (HTML)</label><textarea id="set_header" class="form-control" rows="3">${escapeHtml(SETTINGS.headerHtml)}</textarea></div>
            <div class="col-12"><label class="form-label">تذييل التقرير (HTML)</label><textarea id="set_footer" class="form-control" rows="2">${escapeHtml(SETTINGS.footerHtml)}</textarea></div>
            <div class="col-md-3 form-check"><input id="set_incHeader" class="form-check-input" type="checkbox" ${SETTINGS.includeHeaderOnPrint?'checked':''}><label class="form-check-label">تضمين الرأس</label></div>
            <div class="col-md-3 form-check"><input id="set_incFooter" class="form-check-input" type="checkbox" ${SETTINGS.includeFooterOnPrint?'checked':''}><label class="form-check-label">تضمين العلامة المائية</label></div>
          </div>
        `,
        signatures: `
          <div class="row g-2">
            <div class="col-md-4"><input id="sig1n" class="form-control" placeholder="الاسم 1" value="${escapeHtml(SETTINGS.signatures[0]?.name||'')}"></div>
            <div class="col-md-4"><input id="sig1t" class="form-control" placeholder="الصفة 1" value="${escapeHtml(SETTINGS.signatures[0]?.title||'')}"></div>
            <div class="col-md-4"><input id="sig1p" class="form-control" placeholder="مكان 1" value="${escapeHtml(SETTINGS.signatures[0]?.place||'')}"></div>
            <div class="col-md-4"><input id="sig2n" class="form-control" placeholder="الاسم 2" value="${escapeHtml(SETTINGS.signatures[1]?.name||'')}"></div>
            <div class="col-md-4"><input id="sig2t" class="form-control" placeholder="الصفة 2" value="${escapeHtml(SETTINGS.signatures[1]?.title||'')}"></div>
            <div class="col-md-4"><input id="sig2p" class="form-control" placeholder="مكان 2" value="${escapeHtml(SETTINGS.signatures[1]?.place||'')}"></div>
            <div class="col-md-4"><input id="sig3n" class="form-control" placeholder="الاسم 3" value="${escapeHtml(SETTINGS.signatures[2]?.name||'')}"></div>
            <div class="col-md-4"><input id="sig3t" class="form-control" placeholder="الصفة 3" value="${escapeHtml(SETTINGS.signatures[2]?.title||'')}"></div>
            <div class="col-md-4"><input id="sig3p" class="form-control" placeholder="مكان 3" value="${escapeHtml(SETTINGS.signatures[2]?.place||'')}"></div>
          </div>
        `,
        backup: `
          <div class="row g-2">
            <div class="col-md-6">
              <div class="border rounded p-3 h-100">
                <h6>نسخ احتياطي للإعدادات</h6>
                <p class="small text-muted">يشمل السمات، الخطوط، الشعار، وتفضيلات التقارير.</p>
                <button class="btn btn-outline-primary" onclick="exportSettingsBackup()">تنزيل نسخة الإعدادات</button>
              </div>
            </div>
            <div class="col-md-6">
              <div class="border rounded p-3 h-100">
                <h6>استعادة إعدادات النظام</h6>
                <p class="small text-muted">استيراد ملف JSON يحتوي على إعدادات الواجهة والطباعة.</p>
                <input type="file" class="form-control" accept="application/json" onchange="importSettingsBackup(this.files[0])">
              </div>
            </div>
          </div>
        `
      };
      const saveControls = active === 'backup' ? '' : `<div class="col-12 text-end mt-3"><button class="btn btn-secondary" onclick="loadAll();renderTab('settings')">إلغاء</button> <button class="btn btn-primary" onclick="saveSettingsFromForm()">حفظ</button></div>`;
      return `<div class="card"><div class="card-header">إعدادات النظام</div><div class="card-body">
        ${tabButtons}
        ${sections[active] || ''}
        ${saveControls}
      </div></div>`;
    }
    function saveSettingsFromForm(){
      const previousApiBase = getApiBaseUrl();
      const previousApiSync = !!SETTINGS.apiSyncEnabled;
      const appName = document.getElementById('set_appName');
      if(appName) SETTINGS.appName = appName.value || SETTINGS.appName;
      const appSubtitle = document.getElementById('set_appSubtitle');
      if(appSubtitle) SETTINGS.appSubtitle = appSubtitle.value || '';
      const header = document.getElementById('set_header');
      if(header) SETTINGS.headerHtml = header.value || '';
      const footerField = document.getElementById('set_footer');
      if(footerField) SETTINGS.footerHtml = footerField.value || '';
      const incHeader = document.getElementById('set_incHeader');
      if(incHeader) SETTINGS.includeHeaderOnPrint = !!incHeader.checked;
      const incFooter = document.getElementById('set_incFooter');
      if(incFooter) SETTINGS.includeFooterOnPrint = !!incFooter.checked;
      const auth = document.getElementById('set_auth');
      if(auth) SETTINGS.authEnabled = !!auth.checked;
      const apiBaseUrl = document.getElementById('set_apiBaseUrl');
      if(apiBaseUrl) SETTINGS.apiBaseUrl = apiBaseUrl.value.trim();
      const apiSyncEnabled = document.getElementById('set_apiSync');
      if(apiSyncEnabled) SETTINGS.apiSyncEnabled = !!apiSyncEnabled.checked;
      const theme = document.getElementById('set_theme');
      if(theme) SETTINGS.themeKey = theme.value || 'classic';
      const loginLayout = document.getElementById('set_loginLayout');
      if(loginLayout) SETTINGS.loginLayout = loginLayout.value || 'stacked';
      const appFont = document.getElementById('set_appFont');
      if(appFont) SETTINGS.appFontFamily = appFont.value || SETTINGS.appFontFamily;
      const formFont = document.getElementById('set_formFont');
      if(formFont) SETTINGS.formFontFamily = formFont.value || SETTINGS.formFontFamily;
      const tabFont = document.getElementById('set_tabFont');
      if(tabFont) SETTINGS.tabFontFamily = tabFont.value || SETTINGS.tabFontFamily;
      const sig1n = document.getElementById('sig1n');
      const sig2n = document.getElementById('sig2n');
      const sig3n = document.getElementById('sig3n');
      if(sig1n || sig2n || sig3n){
        SETTINGS.signatures = [
          {name: document.getElementById('sig1n')?.value || '', title: document.getElementById('sig1t')?.value || '', place: document.getElementById('sig1p')?.value || ''},
          {name: document.getElementById('sig2n')?.value || '', title: document.getElementById('sig2t')?.value || '', place: document.getElementById('sig2p')?.value || ''},
          {name: document.getElementById('sig3n')?.value || '', title: document.getElementById('sig3t')?.value || '', place: document.getElementById('sig3p')?.value || ''}
        ];
      }
      const f = document.getElementById('set_logo')?.files?.[0];
      if(f){
        const r = new FileReader();
        r.onload=()=>{
          SETTINGS.logoData = r.result;
          saveAll();
          updateNavLogo();
          showToast('تم حفظ الإعدادات مع الشعار','success');
          renderTab('settings');
        };
        r.readAsDataURL(f);
        return;
      }
     saveAll();
      const nextApiBase = getApiBaseUrl();
      const apiConfigChanged = previousApiBase !== nextApiBase || previousApiSync !== !!SETTINGS.apiSyncEnabled;
      if(apiConfigChanged){
        remoteLoadAttempted = false;
        requestRemoteLoad();
        scheduleRemoteSave();
      }
      document.getElementById('appNameTitle').textContent = SETTINGS.appName;
      document.getElementById('appSubtitle').textContent = SETTINGS.appSubtitle||'';
      const footer = document.getElementById('appFooterText');
      if(footer){
        const year = new Date().getFullYear();
        footer.textContent = `© ${year} ${SETTINGS.appName} — جميع الحقوق محفوظة`;
      }
      applyTheme();
      applyFontSettings();
      updateNavLogo();
      showToast('تم حفظ الإعدادات','success');
      renderTab('settings');
    }
      const footer = document.getElementById('appFooterText');
     
      function renderUsersTable(){
      const u = SETTINGS.users||[];
      if(!u.length) return '<div class="text-muted">لا مستخدمين</div>';
      const rows = u.map((us,i)=>{
        const officer = us.officerId ? (officers.find(o=>o.id===us.officerId)?.name || 'غير مرتبط') : 'غير مرتبط';
        const created = us.createdAt ? new Date(us.createdAt).toLocaleDateString('ar-EG') : '';
        return `<tr>
          <td>${escapeHtml(us.name)}</td>
          <td>${escapeHtml(us.fullName||'')}</td>
          <td>${escapeHtml(us.email||'')}</td>
          <td>${escapeHtml(us.role)}</td>
          <td>${escapeHtml(officer)}</td>
          <td>${escapeHtml(us.phone||'')}</td>
          <td>${escapeHtml(us.note||'')}</td>
          <td>${escapeHtml(created)}</td>
          <td class="text-end text-nowrap"><button class="btn btn-sm btn-outline-secondary me-1" onclick="editUser(${i})">تعديل</button><button class="btn btn-sm btn-outline-warning me-1" onclick="adminResetPasswordForUser(${i})">إعادة ضبط</button><button class="btn btn-sm btn-danger" onclick="removeUserFromSettings(${i})">حذف</button></td>
        </tr>`;
        }).join('');
      return `<div class="table-responsive"><table class="table table-sm table-bordered align-middle"><thead class="table-dark"><tr><th>المستخدم</th><th>الاسم</th><th>البريد</th><th>الدور</th><th>الضابط</th><th>الهاتف</th><th>ملاحظة</th><th>أنشئ</th><th></th></tr></thead><tbody>${rows}</tbody></table></div>`;
    }
    function getSelectedPrivilegeTabs(){
      return Array.from(document.querySelectorAll('.admin-privilege-chk:checked')).map(chk=>chk.value);
    }
    function updatePrivilegeControlState(){
      const role = document.getElementById('admin_user_role')?.value || 'user';
      const container = document.getElementById('adminTabPrivileges');
      if(!container) return;
      const isEditable = role === 'user' || role === 'editor';
      container.querySelectorAll('.admin-privilege-chk').forEach(chk=>{
        chk.disabled = !isEditable;
        if(!isEditable) chk.checked = false;
      });
    }
    function applyUserPrivilegesToForm(role, privileges){
      const selected = Array.isArray(privileges) && privileges.length
        ? privileges
        : (role === 'user' || role === 'editor' ? getRoleTabs(role) : []);
      document.querySelectorAll('.admin-privilege-chk').forEach(chk=>{
        chk.checked = selected.includes(chk.value);
      });
      updatePrivilegeControlState();
    }
    function getViewerUsers(){
      return (SETTINGS.users||[]).filter(u=>u.role==='viewer').map(u=>{
        const officer = officers.find(o=>o.id===u.officerId);
        return {
          username: u.name,
          fullName: u.fullName || officer?.name || '',
          rank: officer?.rank || '',
          password: u.pwd || '',
          officerId: u.officerId || null
        };
      });
    }
    function renderViewerUsersSheet(){
      const viewers = getViewerUsers();
      if(!viewers.length) return '<div class="text-muted">لم يتم إنشاء مستخدمين من نوع المشاهد بعد.</div>';
      const rows = viewers.map((v,i)=>`<tr>
        <td>${i+1}</td>
        <td>${escapeHtml(v.rank)}</td>
        <td>${escapeHtml(v.fullName)}</td>
        <td>${escapeHtml(v.username)}</td>
        <td>${escapeHtml(v.password)}</td>
      </tr>`).join('');
      return `<div class="table-responsive"><table class="table table-sm table-bordered align-middle">
        <thead class="table-dark"><tr><th>#</th><th>الرتبة</th><th>الاسم الكامل</th><th>اسم المستخدم</th><th>كلمة المرور</th></tr></thead>
        <tbody>${rows}</tbody>
      </table></div>`;
    }
    function removeUserFromSettings(i){
      SETTINGS.users.splice(i,1);
      saveAll();
      const usersTable = document.getElementById('admin_usersTable');
      if(usersTable) usersTable.innerHTML = renderUsersTable();
      const viewerSheet = document.getElementById('viewerUsersSheet');
      if(viewerSheet) viewerSheet.innerHTML = renderViewerUsersSheet();
    }
     function editUser(i){
      const u = (SETTINGS.users||[])[i]; if(!u) return;
      document.getElementById('admin_user_name').value = u.name;
      document.getElementById('admin_user_pwd').value = u.pwd;
      document.getElementById('admin_user_role').value = u.role;
      const sel = document.getElementById('admin_user_officer'); if(sel) sel.value = u.officerId || '';
      if(document.getElementById('admin_user_phone')) document.getElementById('admin_user_phone').value = u.phone||'';
      if(document.getElementById('admin_user_note')) document.getElementById('admin_user_note').value = u.note||'';
      if(document.getElementById('admin_user_fullName')) document.getElementById('admin_user_fullName').value = u.fullName||'';
      if(document.getElementById('admin_user_email')) document.getElementById('admin_user_email').value = u.email||'';
      document.getElementById('admin_user_name').dataset.editIndex = i;
      applyUserPrivilegesToForm(u.role, u.tabPrivileges);
    }
    function addUserFromSettings(){}

    function renderSupportRequestsTable(){
      if(!supportRequests.length) return '<div class="text-muted">لا توجد طلبات مسجلة.</div>';
      const rows = supportRequests
        .slice()
        .sort((a,b)=> String(b.createdAt || '').localeCompare(String(a.createdAt || '')))
        .map(r=>{
        const statusCls = r.status==='pending' ? 'warning' : (r.status==='denied'?'danger':'success');
        const created = r.createdAt ? new Date(r.createdAt).toLocaleString('ar-EG') : '';
        const note = escapeHtml(r.note||'');
        return `<tr>
          <td>${escapeHtml(r.username)}</td>
          <td>${escapeHtml(r.type || '')}</td>
          <td><span class="badge bg-${statusCls} text-white">${r.status==='pending'?'بانتظار':(r.status==='denied'?'مرفوض':'مكتمل')}</span></td>
          <td>${created}</td>
          <td>${note}</td>
          <td class="text-end text-nowrap">
            <button class="btn btn-sm btn-outline-primary me-1" onclick="handleSupportRequest(${r.id})" ${r.status==='done'?'disabled':''}>تعيين/إعادة تعيين</button>
            <button class="btn btn-sm btn-outline-secondary" onclick="resolveSupportRequest(${r.id}, 'done')">تمييز كمكتمل</button>
          </td>
        </tr>`;
      }).join('');
      return `<div class="table-responsive"><table class="table table-sm table-bordered align-middle"><thead class="table-dark"><tr><th>المستخدم</th><th>النوع</th><th>الحالة</th><th>التوقيت</th><th>ملاحظات</th><th></th></tr></thead><tbody>${rows}</tbody></table></div>`;
    }
    function refreshSupportRequestsTable(){ const holder = document.getElementById('supportRequestsTable'); if(holder) holder.innerHTML = renderSupportRequestsTable(); }
    function clearResolvedRequests(){ supportRequests = supportRequests.filter(r=>r.status!=='done'); saveAll(); refreshSupportRequestsTable(); }
    function resolveSupportRequest(id, status='done', note=''){
      const req = supportRequests.find(r=>r.id===id);
      if(!req) return;
      req.status = status;
      req.resolvedAt = new Date().toISOString();
      if(note) req.note = `${note} | ${req.note||''}`;
      saveAll();
      refreshSupportRequestsTable();
      logActivity('تحديث طلب مساعدة', {target: req.username, status});
    }
    function handleSupportRequest(id){
      const req = supportRequests.find(r=>r.id===id);
      if(!req){ showToast('لم يتم العثور على الطلب','danger'); return; }
      const idx = (SETTINGS.users||[]).findIndex(u=>u.name===req.username);
      if(idx<0){ showToast('المستخدم غير موجود في النظام','danger'); return; }
      adminResetPasswordForUser(idx, id);
    }

 /* ========= Account ========= */
    function renderAccountTab(){
      if(!SETTINGS.authEnabled) return '<div class="alert alert-info">نظام الدخول غير مفعل حالياً.</div>';
      if(!CURRENT_USER) return signInInlineFormMarkup() + '<div class="alert alert-warning">سجّل الدخول لإدارة حسابك.</div>';
      const u = currentUserRecord();
      if(!u) return '<div class="alert alert-danger">تعذر تحميل بيانات المستخدم الحالي.</div>';
      const officerName = u.officerId ? (officers.find(o=>o.id===u.officerId)?.name || '') : '';
      const created = u.createdAt ? new Date(u.createdAt).toLocaleString('ar-EG') : '';
     const mustChange = !!u.mustChangePassword;
      return `<div class="card">
        <div class="card-header d-flex justify-content-between align-items-center"><div>بيانات حساب المستخدم</div><div class="small text-muted">إنشاء: ${escapeHtml(created)}</div></div>
        <div class="card-body">
          ${mustChange ? '<div class="alert alert-warning">يرجى تغيير كلمة المرور قبل استخدام بقية النظام.</div>' : ''}
          <div class="row g-3 mb-3">
            <div class="col-md-4"><label class="form-label">اسم الدخول (ثابت)</label><input class="form-control" value="${escapeHtml(u.name)}" disabled></div>
            <div class="col-md-4"><label class="form-label">الاسم الكامل</label><input id="acct_fullName" class="form-control" value="${escapeHtml(u.fullName||'')}" placeholder="الاسم المعتمد"></div>
            <div class="col-md-4"><label class="form-label">الدور</label><input class="form-control" value="${escapeHtml(u.role)}" disabled></div>
            <div class="col-md-4"><label class="form-label">البريد الإلكتروني</label><input id="acct_email" type="email" class="form-control" value="${escapeHtml(u.email||'')}" placeholder="you@example.com"></div>
            <div class="col-md-4"><label class="form-label">الهاتف</label><input id="acct_phone" class="form-control" value="${escapeHtml(u.phone||'')}" placeholder="05xxxxxxxx"></div>
            <div class="col-md-4"><label class="form-label">ملاحظة</label><input id="acct_note" class="form-control" value="${escapeHtml(u.note||'')}" placeholder="معلومات إضافية"></div>
            <div class="col-md-4"><label class="form-label">مرتبط بضابط</label><input class="form-control" value="${escapeHtml(officerName || 'غير مرتبط')}" disabled></div>
          </div>
          <div class="d-flex justify-content-between align-items-center mb-4">
            <div class="text-muted small">قم بحفظ أي تغييرات قبل المغادرة.</div>
            <button class="btn btn-success" onclick="saveAccountProfile()">حفظ بيانات الحساب</button>
          </div>
          <hr>
          <h6 class="mb-3">تغيير كلمة المرور</h6>
          <div class="row g-3 mb-2">
            <div class="col-md-4"><input id="acct_pwd_current" type="password" class="form-control" placeholder="كلمة المرور الحالية"></div>
            <div class="col-md-4"><input id="acct_pwd_new" type="password" class="form-control" placeholder="كلمة المرور الجديدة" oninput="updateStrengthBar(this.value,'acct_pwd_strength')"></div>
            <div class="col-md-4"><input id="acct_pwd_confirm" type="password" class="form-control" placeholder="تأكيد كلمة المرور"></div>
          </div>
          <div class="pwd-strength mb-2" id="acct_pwd_strength"><div class="bar"></div></div>
          <div class="text-end">
            <button class="btn btn-primary" onclick="changeOwnPassword()">تحديث كلمة المرور</button>
          </div>
        </div>
      </div>`;
    }

    function saveAccountProfile(){
      const idx = currentUserIndex();
      if(idx<0) return showToast('حساب غير معروف','danger');
      const u = SETTINGS.users[idx];
      u.fullName = document.getElementById('acct_fullName').value.trim();
      u.email = document.getElementById('acct_email').value.trim();
      u.phone = document.getElementById('acct_phone').value.trim();
      u.note = document.getElementById('acct_note').value.trim();
      saveAll();
      logActivity('تحديث ملف مستخدم', {target:u.name});
      showToast('تم حفظ البيانات','success');
    }

    function changeOwnPassword(){
      const idx = currentUserIndex();
      if(idx<0) return showToast('حساب غير معروف','danger');
      const u = SETTINGS.users[idx];
      const cur = document.getElementById('acct_pwd_current').value;
      const np = document.getElementById('acct_pwd_new').value;
      const cp = document.getElementById('acct_pwd_confirm').value;
      if(!cur || !np || !cp) return showToast('أكمل الحقول المطلوبة','danger');
      if(u.pwd !== cur) return showToast('كلمة المرور الحالية غير صحيحة','danger');
      if(np !== cp) return showToast('تأكيد كلمة المرور غير متطابق','danger');
      if(passwordStrengthScore(np) < 2) return showToast('اختر كلمة مرور أقوى (أحرف كبيرة/صغيرة، أرقام، رموز)','danger');
      u.pwd = np;
      u.mustChangePassword = false;
      u.lastLoginAt = new Date().toISOString();
      if(CURRENT_USER) CURRENT_USER.mustChangePassword = false;
      saveAll();
      applyNavPermissions();
      logActivity('تغيير كلمة المرور', {actor: u.name});
      showToast('تم تحديث كلمة المرور','success');
      document.getElementById('acct_pwd_current').value='';
      document.getElementById('acct_pwd_new').value='';
      document.getElementById('acct_pwd_confirm').value='';
      renderTab('account');
    }

    /* ========= About ========= */
     function renderAboutTab(){
      return `<div class="card">
        <div class="card-header">حول التطبيق</div>
        <div class="card-body">
          <p class="mb-3">منصة تنظيم وإدارة خدمات الإدارة العامة للمساعدات الفنية، مصممة لتقديم تجربة موثوقة ومتكاملة لإدارة الجداول، الضباط، والعمليات الداعمة.</p>
          <div class="row g-2">
            <div class="col-md-6">
              <h6 class="mb-2">قدرات تشغيلية</h6>
              <ul>
                <li>توليد جداول شهرية ذكية تراعي الرتب، أوقات الراحة، والاستثناءات.</li>
                <li>تقارير جاهزة للطباعة A4 مع تواقيع وعلامة مائية.</li>
                <li>أرشفة شهرية واستعراض إحصاءات الخدمة حسب الضباط.</li>
              </ul>
            </div>
            <div class="col-md-6">
              <h6 class="mb-2">حوكمة وإدارة</h6>
              <ul>
                <li>لوحة تحكم لنسخ احتياطي واستعادة البيانات.</li>
                <li>صلاحيات متعددة للمستخدمين مع سجل نشاط شامل.</li>
                <li>تخصيص السمات والخطوط والشعار بما يتناسب مع الهوية المؤسسية.</li>
              </ul>
            </div>
          </div>
          <div class="mt-3"><strong>الإصدار:</strong> 1.0 — إصدار تشغيلي موحد</div>
        </div>
      </div>`;
    }

   /* ========= Init ========= */
 loadAll();
    startSplashProgress();
