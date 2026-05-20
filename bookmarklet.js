javascript: (function () {
    if (!location.hostname.includes('17.live')) {
        alert('このBookmarkletは 17.live でのみ動作します。\n現在のホスト: ' + location.hostname);
        return;
    }

    // ===== localStorage から認証情報を取得 =====
    let loginInfo;
    try {
        const raw = localStorage.getItem('17LIVE/LOGIN_INFO');
        if (!raw) {
            alert('ログイン情報が見つかりません。17.live にログインしてから再度実行してください。');
            return;
        }
        loginInfo = JSON.parse(raw);
    } catch (e) {
        alert('ログイン情報の読み取りに失敗しました: ' + e.message);
        return;
    }

    const ACCESS_TOKEN = loginInfo.accessToken;
    const BEARER_RAW = loginInfo.jwtAccessToken;
    if (!ACCESS_TOKEN || !BEARER_RAW) {
        alert('accessToken または jwtAccessToken が取得できませんでした。\nログインし直してください。');
        return;
    }
    const BEARER = /^Bearer\s+/i.test(BEARER_RAW) ? BEARER_RAW : 'Bearer ' + BEARER_RAW;

    // ===== JWT から userID を抽出 =====
    let userID;
    try {
        const t = BEARER.replace(/^Bearer\s+/i, '');
        const parts = t.split('.');
        if (parts.length !== 3) throw new Error('JWT 形式不正');
        const payload = JSON.parse(atob(parts[1].replace(/-/g, '+').replace(/_/g, '/')));
        userID = payload.userID;
        if (!userID) throw new Error('userID が JWT に含まれていません');
    } catch (e) {
        alert('JWT 解析失敗: ' + e.message);
        return;
    }

    const APP_VERSION = "3.328.0";
    const ACTION = "getReceivedGiftLog";
    const PAGE_SIZE = 10000;
    const MAX_PAGES = 1000;

    // ===== ローディングオーバーレイ =====
    const overlay = document.createElement('div');
    overlay.style.cssText = 'position:fixed;top:0;left:0;width:100vw;height:100vh;background:rgba(0,0,0,0.75);z-index:2147483647;display:flex;align-items:center;justify-content:center;font-family:sans-serif;color:#fff;';
    const card = document.createElement('div');
    card.style.cssText = 'background:#1a1a1a;padding:32px 40px;border-radius:12px;text-align:center;min-width:360px;box-shadow:0 8px 32px rgba(0,0,0,0.5);';
    const spinner = document.createElement('div');
    spinner.style.cssText = 'width:48px;height:48px;border:4px solid #333;border-top-color:#ff5722;border-radius:50%;margin:0 auto 20px;animation:__giftSpin 1s linear infinite;';
    const style = document.createElement('style');
    style.textContent = '@keyframes __giftSpin{to{transform:rotate(360deg)}}';
    document.head.appendChild(style);
    const msg = document.createElement('div');
    msg.style.cssText = 'font-size:16px;font-weight:bold;margin-bottom:8px;';
    msg.textContent = '初期化中...';
    const sub = document.createElement('div');
    sub.style.cssText = 'font-size:13px;color:#aaa;line-height:1.5;';
    card.appendChild(spinner);
    card.appendChild(msg);
    card.appendChild(sub);
    overlay.appendChild(card);
    document.body.appendChild(overlay);
    const setStatus = function (main, detail) {
        msg.textContent = main;
        sub.textContent = detail || '';
        console.log('[Status]', main, detail || '');
    };
    const closeOverlay = function () { if (overlay.parentNode) overlay.parentNode.removeChild(overlay); };
    const errorOverlay = function (text) {
        spinner.style.display = 'none';
        msg.textContent = '❌ エラー';
        sub.textContent = text;
        sub.style.color = '#ff6b6b';
        setTimeout(closeOverlay, 5000);
    };

    const randomDeviceID = function () {
        const chars = 'abcdefghijklmnopqrstuvwxyz0123456789';
        const bytes = new Uint8Array(32);
        crypto.getRandomValues(bytes);
        let r = '';
        for (let i = 0; i < 32; i++) r += chars[bytes[i] % chars.length];
        return r;
    };

    const loadXLSX = function () {
        return new Promise((resolve, reject) => {
            if (window.XLSX) { resolve(window.XLSX); return; }
            const s = document.createElement('script');
            s.src = 'https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js';
            s.onload = () => resolve(window.XLSX);
            s.onerror = () => reject(new Error('SheetJS 読み込み失敗'));
            document.head.appendChild(s);
        });
    };

    // ===== 1ページ取得 =====
    const fetchPage = function (beforeTime) {
        const body = {
            key: "",
            data: JSON.stringify({
                userID: userID,
                version: APP_VERSION,
                accessToken: ACCESS_TOKEN,
                nonce: crypto.randomUUID().toUpperCase(),
                count: String(PAGE_SIZE),
                ipCountry: "JP",
                deviceType: "IOS",
                packageName: "com.machipopo.story17",
                action: ACTION,
                Authorization: BEARER,
                language: "JP",
                OSVersion: "26.5",
                deviceID: randomDeviceID(),
                hardware: "iPhone18,2",
                beforeTime: String(beforeTime)
            }),
            cypher: "0_v2"
        };
        return fetch('https://wap-api.17app.co/apiGateWay', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(body)
        })
            .then(r => r.text().then(t => ({ status: r.status, body: t })))
            .then(res => {
                if (res.status !== 200) throw new Error('status=' + res.status);
                const outer = JSON.parse(res.body);
                const inner = JSON.parse(outer.data);
                if (!Array.isArray(inner)) throw new Error('not array');
                return inner;
            });
    };

    // ===== 全件取得(ページネーション) =====
    const fetchAll = async function () {
        const all = [];
        let beforeTime = 2147483647;
        let pageNum = 0;
        const t0 = Date.now();

        while (pageNum < MAX_PAGES) {
            pageNum++;
            const elapsedBefore = ((Date.now() - t0) / 1000).toFixed(1);
            setStatus('ギフト履歴を取得中...', 'ページ ' + pageNum + ' を取得中 / 累計 ' + all.length + ' 件 / 経過 ' + elapsedBefore + ' 秒');

            let page;
            try {
                page = await fetchPage(beforeTime);
            } catch (e) {
                console.error('[fetchPage] error', e);
                throw new Error('ページ ' + pageNum + ' で失敗: ' + e.message);
            }

            const elapsedAfter = ((Date.now() - t0) / 1000).toFixed(1);
            setStatus('ギフト履歴を取得中...', 'ページ ' + pageNum + ': ' + page.length + ' 件取得 / 累計 ' + (all.length + page.length) + ' 件 / 経過 ' + elapsedAfter + ' 秒');

            console.log('[page ' + pageNum + ']', page.length + '件', 'beforeTime=' + beforeTime);

            if (page.length === 0) {
                console.log('[完了] 空ページ到達');
                break;
            }

            for (let i = 0; i < page.length; i++) all.push(page[i]);

            if (page.length < PAGE_SIZE) {
                console.log('[完了] PAGE_SIZE 未満のページ到達');
                break;
            }

            let minTs = Infinity;
            for (let i = 0; i < page.length; i++) {
                const ts = page[i].timestamp;
                if (ts < minTs) minTs = ts;
            }
            if (minTs >= beforeTime) {
                console.log('[完了] beforeTime が更新されなかった(無限ループ防止)');
                break;
            }
            beforeTime = minTs;
        }

        return all;
    };

    // ===== Excel 出力 =====
    const exportToExcel = function (data) {
        setStatus('Excel 生成中...', 'SheetJS ライブラリを読み込んでいます');
        return loadXLSX().then(XLSX => {
            setStatus('Excel 生成中...', '全履歴シート(' + data.length + '件)');
            const rows = data.map((r, i) => ({
                '#': i + 1,
                '日時': new Date(r.timestamp * 1000).toLocaleString('ja-JP'),
                'タイムスタンプ': r.timestamp,
                '贈り主表示名': r.userInfo?.displayName || '',
                'openID': r.userInfo?.openID || '',
                '本名/コメント': r.userInfo?.name || '',
                '贈り主userID': r.userInfo?.userID || '',
                'roomID': r.userInfo?.roomID || '',
                'level': r.userInfo?.level || '',
                'baller': r.userInfo?.baller || 0,
                'ballerLevel': r.userInfo?.ballerLevel || 0,
                'region': r.userInfo?.region || '',
                'ギフト名': r.giftInfo?.name || '',
                'giftID': r.giftInfo?.giftID || '',
                'ポイント': r.giftInfo?.point || 0,
                'キャンセル': r.isCanceled || 0,
                'website': r.userInfo?.website || '',
                'bio': (r.userInfo?.bio || '').replace(/\n/g, ' ')
            }));

            setStatus('Excel 生成中...', '贈り主集計中');
            const bySender = {};
            data.forEach(r => {
                const id = r.userInfo?.userID || '?';
                if (!bySender[id]) {
                    bySender[id] = {
                        userID: id,
                        displayName: r.userInfo?.displayName || '',
                        openID: r.userInfo?.openID || '',
                        件数: 0, 合計ポイント: 0,
                        最古: r.timestamp, 最新: r.timestamp,
                        level: r.userInfo?.level || 0,
                        baller: r.userInfo?.baller || 0
                    };
                }
                const s = bySender[id];
                s.件数++;
                s.合計ポイント += (r.giftInfo?.point || 0);
                if (r.timestamp < s.最古) s.最古 = r.timestamp;
                if (r.timestamp > s.最新) s.最新 = r.timestamp;
            });
            const ranking = Object.values(bySender)
                .sort((a, b) => b.合計ポイント - a.合計ポイント)
                .map((s, i) => ({
                    順位: i + 1,
                    displayName: s.displayName, openID: s.openID, userID: s.userID,
                    件数: s.件数, 合計ポイント: s.合計ポイント,
                    平均ポイント: Math.round(s.合計ポイント / s.件数 * 100) / 100,
                    最古: new Date(s.最古 * 1000).toLocaleString('ja-JP'),
                    最新: new Date(s.最新 * 1000).toLocaleString('ja-JP'),
                    level: s.level, baller: s.baller
                }));

            setStatus('Excel 生成中...', 'ギフト集計中');
            const byGift = {};
            data.forEach(r => {
                const g = r.giftInfo?.giftID || '?';
                if (!byGift[g]) {
                    byGift[g] = {
                        giftID: g,
                        ギフト名: r.giftInfo?.name || '',
                        単価ポイント: r.giftInfo?.point || 0,
                        件数: 0, 合計ポイント: 0
                    };
                }
                byGift[g].件数++;
                byGift[g].合計ポイント += (r.giftInfo?.point || 0);
            });
            const giftRanking = Object.values(byGift)
                .sort((a, b) => b.合計ポイント - a.合計ポイント)
                .map((g, i) => ({ 順位: i + 1, ...g }));

            let oldestTs = Infinity, newestTs = -Infinity;
            for (let i = 0; i < data.length; i++) {
                const ts = data[i].timestamp;
                if (ts < oldestTs) oldestTs = ts;
                if (ts > newestTs) newestTs = ts;
            }
            const totalPt = data.reduce((a, r) => a + (r.giftInfo?.point || 0), 0);
            const summary = [
                { 項目: '対象userID', 値: userID },
                { 項目: '総件数', 値: data.length },
                { 項目: '総ポイント', 値: totalPt },
                { 項目: '贈り主ユニーク数', 値: Object.keys(bySender).length },
                { 項目: 'ギフト種類数', 値: Object.keys(byGift).length },
                { 項目: '最古', 値: data.length ? new Date(oldestTs * 1000).toLocaleString('ja-JP') : '-' },
                { 項目: '最新', 値: data.length ? new Date(newestTs * 1000).toLocaleString('ja-JP') : '-' },
                { 項目: '生成日時', 値: new Date().toLocaleString('ja-JP') }
            ];

            setStatus('Excel 生成中...', '書き出し中');
            const wb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(summary), 'サマリ');
            XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(ranking), '贈り主ランキング');
            XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(giftRanking), 'ギフトランキング');
            XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows), '全履歴');

            const filename = 'gift_log_' + userID.substring(0, 8) + '_' + new Date().toISOString().replace(/[:.]/g, '-').substring(0, 19) + '.xlsx';
            XLSX.writeFile(wb, filename);

            return { filename, totalPt, senderCount: Object.keys(bySender).length, giftCount: Object.keys(byGift).length };
        });
    };

    // ===== メインフロー =====
    (async function () {
        setStatus('認証情報を確認しました', 'userID: ' + userID.substring(0, 8) + '...');
        console.log('[userID]', userID);
        console.log('[accessToken]', ACCESS_TOKEN);

        const t0 = Date.now();
        let all;
        try {
            all = await fetchAll();
        } catch (e) {
            errorOverlay('取得失敗: ' + e.message);
            return;
        }
        const fetchElapsed = ((Date.now() - t0) / 1000).toFixed(1);
        console.log('[全件取得] 件数=' + all.length + ' / ' + fetchElapsed + '秒');

        if (all.length === 0) {
            errorOverlay('データが0件でした');
            return;
        }

        window.__lastResult = all;

        try {
            const r = await exportToExcel(all);
            closeOverlay();
            alert('完了\n\n取得時間: ' + fetchElapsed + '秒 / ' + all.length + '件\n\nファイル: ' + r.filename + '\n贈り主: ' + r.senderCount + '名\nギフト種: ' + r.giftCount + '種\n総ポイント: ' + r.totalPt);
        } catch (e) {
            errorOverlay('Excel 生成失敗: ' + e.message);
        }
    })();
})();