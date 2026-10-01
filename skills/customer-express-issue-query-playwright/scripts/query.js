(async ({ nos, timeoutMs = 30000 } = {}) => {
  // Evaluated by Playwright in the selected, already logged-in order frame.
  // Only the required four result fields and aggregate counts leave this frame.
  const logs = [];
  const clean = value => value == null ? '' : String(value).replace(/\s+/g, ' ').trim();
  const norm = value => clean(value).replace(/^@/, '').toUpperCase();
  class QueryError extends Error {}

  try {
    if (!Array.isArray(nos) || !nos.length || nos.length > 60 ||
        nos.some(no => typeof no !== 'string' || !no.trim() || !norm(no))) {
      throw new QueryError('Invalid tracking batch');
    }
    if (!Number.isFinite(timeoutMs) || timeoutMs < 1 || timeoutMs > 300000) {
      throw new QueryError('Invalid request timeout');
    }
    if (window.location.pathname !== '/app/order/order/list.aspx' ||
        !document.getElementById('__VIEWSTATE')) {
      throw new QueryError('Open the high edition order page first');
    }

    const matches = new Map(nos.map(no => [no, []]));
    const fields = ['l_id', 'plat_l_id', 'purchase_lid', 'multiWaybillLid',
      'logistics_no', 'express_no', 'waybillNo', 'waybill_no', 'lc_id'];

    async function request(batch) {
      const hidden = id => clean(document.getElementById(id)?.value);
      const tokens = [...new Set(batch.flatMap(no =>
        [no, no.startsWith('@') ? no.slice(1) : '@' + no]))];
      const form = {
        __VIEWSTATE: hidden('__VIEWSTATE'),
        __VIEWSTATEGENERATOR: hidden('__VIEWSTATEGENERATOR'),
        insurePrice: hidden('hdInsurePrice'),
        _jt_page_count_enabled: '', _jt_page_increament_enabled: 'true',
        _jt_page_increament_page_mode: '', _jt_page_increament_key_value: '',
        _jt_page_increament_business_values: '', _jt_page_increament_key_name: 'o_id',
        _jt_page_size: 500, _jt_page_action: '1', fe_node_desc: '',
        receiver_state: '', receiver_city: '', receiver_district: '', receiver_address: '',
        receiver_name: '', receiver_phone: '', receiver_mobile: '', check_name: '',
        check_address: '', fe_remark_type: 'single', node_type: '', fe_flag: '',
        fe_is_append_remark: '', __CALLBACKID: 'JTable1',
        __CALLBACKPARAM: JSON.stringify({
          Method: 'LoadDataToJSON',
          Args: ['1', JSON.stringify([{ k: 'l_id', v: tokens.join(','), c: '@=' }]), '{}']
        })
      };
      const controller = new AbortController();
      let timedOut = false;
      const timer = setTimeout(() => {
        timedOut = true;
        controller.abort();
      }, timeoutMs);
      let responseText;
      try {
        const response = await window.fetch(
          window.location.origin + '/app/order/order/list.aspx?_c=jst-epaas&epaas=true&ts___=' +
            Date.now() + '&am___=LoadDataToJSON',
          {
            method: 'POST', credentials: 'same-origin', mode: 'same-origin', redirect: 'error',
            headers: { 'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8', Accept: '*/*' },
            body: new URLSearchParams(form).toString(), signal: controller.signal
          }
        );
        if (response.status !== 200) {
          const status = Number.isInteger(response.status) ? response.status : 'error';
          throw new QueryError('HTTP ' + status + '; check browser login or permissions');
        }
        responseText = await response.text();
      } catch (error) {
        if (timedOut) throw new QueryError('Order request timed out; check the browser and retry');
        if (error instanceof QueryError) throw error;
        throw new QueryError('Order request failed; check browser login, verification, or connection');
      } finally {
        clearTimeout(timer);
      }

      let outer;
      try {
        // ASP.NET callbacks may prefix JSON with a numeric length and a pipe.
        // A pipe inside an ordinary JSON remark must never be treated as a prefix.
        outer = JSON.parse(responseText.replace(/^\s*\d+\|/, ''));
      } catch (_) {
        throw new QueryError('Invalid response; check login or verification in browser');
      }
      if (!outer || outer.IsSuccess !== true) {
        throw new QueryError('API unsuccessful; inspect browser permissions or verification');
      }
      let value;
      try {
        value = typeof outer.ReturnValue === 'string'
          ? JSON.parse(outer.ReturnValue || '{}') : outer.ReturnValue;
      } catch (_) {
        throw new QueryError('Unexpected order response schema');
      }
      const rows = value?.datas || value?.data;
      if (!Array.isArray(rows) || rows.some(row => !row || typeof row !== 'object' || Array.isArray(row))) {
        throw new QueryError('Unexpected order response schema');
      }
      if (rows.length >= 500) {
        logs.push({ requested: batch.length, rows: rows.length, matched: 0, fallback: 0, split: batch.length > 1 });
        if (batch.length <= 1) {
          throw new QueryError('Single tracking number reaches page limit; cannot certify completeness');
        }
        const half = Math.ceil(batch.length / 2);
        await request(batch.slice(0, half));
        await request(batch.slice(half));
        return;
      }

      let fallback = 0;
      for (const row of rows) {
        const keys = new Set(fields.flatMap(field => clean(row[field])
          .split(/[，、;,|\s]+/).map(norm)).filter(Boolean));
        let found = batch.filter(no => keys.has(norm(no)));
        if (!found.length) {
          const blob = JSON.stringify(row).toUpperCase();
          found = batch.filter(no => {
            const escaped = norm(no).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
            return new RegExp('(^|[^A-Z0-9])@?' + escaped + '([^A-Z0-9]|$)').test(blob);
          });
          fallback += found.length;
        }
        const key = JSON.stringify([
          row.o_id ?? row.so_id ?? row.raw_so_id ?? row.pre_so_id,
          row.confirm_date, row.paid_amount, row.remark, row.l_id,
          row.plat_l_id, row.purchase_lid, row.multiWaybillLid
        ].map(clean));
        for (const no of found) {
          if (!matches.get(no).some(match => match.key === key)) {
            matches.get(no).push({ key, data: [no, clean(row.confirm_date), clean(row.paid_amount), clean(row.remark)] });
          }
        }
      }
      logs.push({ requested: batch.length, rows: rows.length,
        matched: batch.filter(no => matches.get(no).length).length, fallback });
    }

    await request(nos);
    return {
      ok: true,
      results: nos.map(no => ({ no, matched: matches.get(no).length,
        data: matches.get(no).length ? matches.get(no).map(match => match.data) : [[no, '', '', '']] })),
      logs
    };
  } catch (error) {
    // Never propagate arbitrary exception messages: they can contain page data or credentials.
    return { ok: false, reason: error instanceof QueryError ? error.message : 'Order query failed; inspect the browser and retry', logs };
  }
})
