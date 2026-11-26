// 선택과목 대시보드_총괄 · 실데이터 어댑터
// 2023~2026: [나이스]학교별학년별학기별개설과목현황(고) 실데이터 (분반 = 행 수, 창의적 체험활동 제외)
// 수강학생: [나이스]기초시간표 확보 학교만 실측 표시 · 미확보 학교는 값 없음(추정하지 않음)
// 교과(군)·과목유형: 교육과정_과목분류_적재기준표.xlsx (경기도 교육과정 총론 2021-486호 · 2024-541호) · 적용 교육과정 + 정규화 과목명 일치만 연결
//   선택과목 = 공통과목을 제외한 개설 과목 (보통교과 일반·진로·융합 + 전문교과 + 기타) · 기준표에 없는 과목은 교과(군) '기타' · 기준표에 없는 과목은 '기준표 외'로 집계에서 제외
// 교원 시수 = 교육활동 교과담당시간(EDAC_SJMT_TKCG_HR, 창체 제외) ÷ 교원수 · 학교 전체 교원수·시수만 사용 (담당과목명 → 교과(군) 연결 코드표 미확보로 교과(군)별 분리 안 함)
// 교원: [교육통계]교원의주당수업시간(전체교원)(상세과목) 2023~2026 실데이터 · 2022는 가상
// 2022 개설 현황만 2023 실데이터를 바탕으로 만든 가상 데이터
(function () {
  function apply() {
  const R = window.CGRAW, V = window.V9D;
  if (!R || !V || !V.CAT) return false;
  if (window.CG_READY && V.schools && V.schools().length && V.schools()[0].code && V.schools()[0].code[0] === 'J') return true;
  window.CG_REAL_FROM = 2023;
  const H = s => { let x = 2166136261; for (let i = 0; i < s.length; i++) { x ^= s.charCodeAt(i); x = Math.imul(x, 16777619); } return (x >>> 0) / 4294967296; };
  const AREAS = V.AREAS.filter(x => x !== '전문교과' && x !== '기타').concat(['전문교과', '기타']);
  const TY = { C: '공통', G: '일반', J: '진로', F: '융합', S: '전문' };
  const NZ = s => (s || '').normalize('NFKC').toUpperCase().replace(/\s+/g, '').replace(/[·⋅･・•∙ㆍ]/g, '·');
  const _rc = {}; const refOf = (name, cu) => { const k = cu + '|' + name; if (k in _rc) return _rc[k]; const r = R.ref && R.ref[cu] && R.ref[cu][NZ(name)]; return (_rc[k] = r ? { area: r[0], type: TY[r[1]] || '전문', disp: !!r[2], id: r[3] } : null); };
  const N15 = new Set(V.CAT.filter(e => e.n15).map(e => e.n15.replace(/\s+/g, '')));
  const COMMON15 = [[/^공통국어/, '국어'], [/^공통수학/, '수학'], [/^공통영어/, '영어'], [/^한국사/, '한국사'], [/^통합사회/, '통합사회'], [/^통합과학/, '통합과학'], [/^과학탐구실험/, '과학탐구실험']];
  const SC = R.schools.map(([code, name, sg, type, found, cpg, students, real]) => ({ code, name, sigungu: sg, office: V.OFFICE_OF[sg], urban: V.URBAN_OF[sg], type, found, cpg: Math.max(1, cpg), students, realStu: !!real, coed: '' }));
  const byCode = {}; SC.forEach(s => { byCode[s.code] = s; });
  const _real = {};
  function realOffers(s, y, sem) {
    const k = s.code + '|' + y + '|' + sem; if (_real[k]) return _real[k];
    const raw = R.off[k], out = [];
    if (raw) raw.split(';').forEach(e => {
      const p = e.split(','), ix = +p[0], g = +p[1], cls = +p[2], stu = +p[3], re = p[4] === '1', h = +p[6] || 0;
      const sj = R.subj[ix], cu = V.cur(y, g), rf = refOf(sj[0], cu), area = rf ? rf.area : '기타';
      const hrs = h || NaN;
      out.push({ e: null, eid: ix, g, cu, name: sj[0], type: rf ? rf.type : '기타', students: re ? stu : NaN, classes: cls, credits: hrs, hours: cls * hrs, dept: area, area, selective: rf ? rf.disp : true, special: !!(rf && rf.type === '전문'), est: !re, inCat: !!rf, refId: rf ? rf.id : null });
    });
    return (_real[k] = out);
  }
  const _syn = {};
  function synOffers(s, y, sem) {
    const k = s.code + '|' + y + '|' + sem; if (_syn[k]) return _syn[k];
    const base = realOffers(s, 2023, sem), out = [];
    base.forEach(o => {
      if (o.selective && H(s.code + '|' + o.name + '|' + o.g + '|' + y) > 0.93) return;
      const cls = Math.max(1, Math.round(o.classes * (o.selective ? 0.97 : 1)));
      out.push(Object.assign({}, o, { students: NaN, classes: cls, hours: cls * o.credits, est: true, syn: true }));
    });
    return (_syn[k] = out);
  }
  const _t = {};
  function teachers(s, y) {
    const k = s.code + '|' + y; if (_t[k]) return _t[k];
    const all = R.tch[s.code] || {}, src = all[y] || (y === 2022 ? all[2023] : null), out = [];
    if (src) Object.entries(src).forEach(([d, [n, h]]) => {
      const hh = (y === 2022 && !all[y] ? h * 0.98 : h) / n;
      for (let i = 0; i < n; i++) out.push({ id: s.code + d + i, dept: '__all', area: '__all', hours: hh, major: hh, non: 0, subj: {}, nSubj: 0, multi: false, solo: false });
    });
    return (_t[k] = out);
  }
  const ss = (s, y, sem) => ({ offers: y >= 2023 ? realOffers(s, y, sem) : synOffers(s, y, sem), teachers: teachers(s, y) });
  const COMMON = new Set(["국어","공통국어1","공통국어2","수학","공통수학1","공통수학2","영어","공통영어1","공통영어2","통합사회","통합사회1","통합사회2","한국사","한국사1","한국사2","통합과학","통합과학1","통합과학2","과학탐구실험","과학탐구실험1","과학탐구실험2","기본수학1","기본수학2","기본영어1","기본영어2"]);
  const EXCL = new Set(['진로활동', '자율활동', '동아리활동', '봉사활동', '창의적 체험활동', '창의적체험활동']);
  const _tt = {};
  const ttOf = (s, y, sem) => { const k = s.code + '|' + y + '|' + sem; if (k in _tt) return _tt[k]; const raw = R.tt && R.tt[k]; if (!raw) return (_tt[k] = null);
    return (_tt[k] = raw.split(';').map(e => { const p = e.split(','), sj = R.subj[+p[0]], g = +p[1], rf = refOf(sj[0], V.cur(y, g)); return { name: sj[0], area: rf ? rf.area : '기타', g, ch: +p[2], sh: +p[3], common: rf ? !rf.disp : false, inCat: !!rf }; }).filter(o => !EXCL.has(o.name))); };
  const enrolled = (s, y) => { let t = 0, c = 0; [1, 2, 3].forEach(g => { const v = R.grT && R.grT[s.code + '|' + y + '|1|' + g]; if (v) { t += v; c++; } }); return c ? { n: t, grades: c } : null; };
  const tchSubj = (s, y) => { const o = (R.tchs && R.tchs[s.code] && R.tchs[s.code][y]) || null; return o ? Object.entries(o).map(([name, [n, h]]) => ({ name, n, h })) : null; };
  const REFCAT = (R.refcat || []).map(([cu, area, t, name]) => ({ cu, area, t: TY[t], name }));
  window.V9D = Object.assign({}, V, { REFCAT, refOf, ttOf, enrolled, tchSubj, AREAS, schools: () => SC, ss, PPC: R.ppc, REAL_STU_N: SC.filter(s => s.realStu).length, IS_REAL: true });
  window.CG_READY = true;
  return true;
  }
  let tries = 0;
  const tick = () => { if (!apply() && tries++ < 600) setTimeout(tick, 50); };
  tick();
  // v9-data.js가 늦게 다시 실행돼 가상 데이터로 덮어쓰는 경우 대비
  setInterval(() => { if (window.V9D && !window.V9D.IS_REAL) { window.CG_READY = false; apply(); } }, 300);
})();
