document.addEventListener('DOMContentLoaded', () => {
  // ---- Conflict policies: must match ConflictPolicy enum in src/cs/ExcelDiffEngine.cs ----
  const ConflictPolicy = Object.freeze({ LastWins: 0, FirstWins: 1, PromptUser: 2 });
  const POLICY_NAMES = ['LastWins', 'FirstWins', 'PromptUser'];

  let basePath = '';
  // { path, checked, selected } — mirrors the desktop CheckedListBox (check = merge target, select = delete target)
  let modFiles = [];
  let isBusy = false;

  const txtBasePath = document.getElementById('txtBasePath');
  const fileListContainer = document.getElementById('fileListContainer');
  const emptyPlaceholder = document.getElementById('emptyPlaceholder');
  const fileCountBadge = document.getElementById('fileCountBadge');
  const terminalBody = document.getElementById('terminalBody');
  const cboConflictPolicy = document.getElementById('cboConflictPolicy');
  const actionButtons = ['btnRun', 'btnReset', 'btnClearBoth'].map(id => document.getElementById(id));

  // ---- Debug log (same role as DebugLog.cs: every terminal line + extra DEBUG/ERROR detail, kept in memory) ----
  const debugLines = [];
  function debugLog(level, msg, detail) {
    const stamp = new Date().toISOString();
    const line = `${stamp} [${level}] ${msg}`;
    debugLines.push(detail ? `${line}\n${detail}` : line);
    if (debugLines.length > 5000) debugLines.shift();
    const fn = level === 'ERROR' ? console.error : level === 'WARN' ? console.warn : level === 'DEBUG' ? console.debug : console.log;
    fn.call(console, `[MergeBooster] ${line}`, detail || '');
  }
  // Call MergeBooster.downloadDebugLog() in DevTools to save the log as a file
  window.MergeBooster = {
    debugLines,
    downloadDebugLog() {
      const blob = new Blob([debugLines.join('\n')], { type: 'text/plain;charset=utf-8' });
      const a = document.createElement('a');
      const d = new Date();
      const ymd = `${d.getFullYear()}${String(d.getMonth() + 1).padStart(2, '0')}${String(d.getDate()).padStart(2, '0')}`;
      a.href = URL.createObjectURL(blob);
      a.download = `MergeBooster_debug_${ymd}.log`;
      a.click();
      URL.revokeObjectURL(a.href);
    }
  };
  window.addEventListener('error', e => debugLog('ERROR', `처리되지 않은 예외: ${e.message}`, e.error && e.error.stack));
  window.addEventListener('unhandledrejection', e => debugLog('ERROR', '처리되지 않은 Promise 거부', String(e.reason && (e.reason.stack || e.reason))));
  debugLog('INFO', `===== MergeBooster(Web) 시작 · ${navigator.userAgent}`);

  function appendLog(msg, color = '#89dceb') {
    const p = document.createElement('p');
    const time = new Date().toLocaleTimeString('ko-KR', { hour12: false });
    p.style.color = color;
    p.textContent = `[${time}] ${msg}`;
    terminalBody.appendChild(p);
    terminalBody.scrollTop = terminalBody.scrollHeight;
    debugLog('LOG', msg);
  }

  function setBusy(busy) {
    isBusy = busy;
    actionButtons.forEach(b => { b.disabled = busy; b.style.opacity = busy ? '0.55' : ''; b.style.cursor = busy ? 'not-allowed' : ''; });
  }

  function fileName(p) {
    return p.split(/[\\/]/).pop();
  }

  // Base File Selection
  document.getElementById('btnBrowseBase').addEventListener('click', () => {
    basePath = 'C:\\Projects\\ExcelData\\연령별인구현황_원본.xlsx';
    txtBasePath.value = basePath;
    appendLog(`원본 파일 지정 완료: ${basePath}`, '#a6e3a1');
  });

  document.getElementById('btnOpenBase').addEventListener('click', () => {
    if (!txtBasePath.value) {
      alert('원본 파일 경로를 먼저 지정해 주세요.');
      return;
    }
    appendLog(`📖 원본 엑셀 파일 실행: ${txtBasePath.value}`, '#89b4fa');
  });

  // Render Modified Files List (DOM API only — file names never go through innerHTML)
  function renderList() {
    fileListContainer.replaceChildren();
    if (modFiles.length === 0) {
      fileListContainer.appendChild(emptyPlaceholder);
      fileCountBadge.textContent = '0개 파일 선택됨';
      return;
    }
    const checkedCount = modFiles.filter(f => f.checked).length;
    fileCountBadge.textContent = `${checkedCount}/${modFiles.length}개 파일 선택됨`;

    modFiles.forEach((file, idx) => {
      const row = document.createElement('div');
      row.style.cssText = 'display: flex; align-items: center; gap: 0.5rem; padding: 0.5rem 0.75rem; border-radius: 8px; font-size: 13px; font-weight: 500; color: #334155; cursor: pointer;' +
        (file.selected ? 'background-color: #eef2ff; border: 1px solid #a5b4fc;' : 'background-color: #f8fafc; border: 1px solid #e2e8f0;');

      const chk = document.createElement('input');
      chk.type = 'checkbox';
      chk.checked = file.checked;
      chk.title = '취합 대상';
      chk.style.cssText = '-webkit-appearance: checkbox; appearance: auto; width: 16px; height: 16px; flex: none; cursor: pointer;';
      chk.addEventListener('click', e => e.stopPropagation());
      chk.addEventListener('change', () => {
        file.checked = chk.checked;
        debugLog('DEBUG', `체크 변경: ${file.path} → ${file.checked}`);
        renderList();
      });

      const label = document.createElement('span');
      label.style.cssText = 'flex: 1; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;';
      label.textContent = `📄 ${fileName(file.path)}`;
      label.title = file.path;

      const del = document.createElement('button');
      del.type = 'button';
      del.textContent = '✕';
      del.title = '목록에서 제거';
      del.style.cssText = 'color: #ef4444; font-weight: 700; cursor: pointer; padding: 0.2rem 0.5rem; border-radius: 4px;';
      del.addEventListener('click', e => {
        e.stopPropagation();
        const [removed] = modFiles.splice(idx, 1);
        appendLog(`목록에서 제거됨: ${removed.path}`, '#f9e2af');
        renderList();
      });

      // Clicking the row toggles its selection for [선택 삭제], like the desktop list
      row.addEventListener('click', () => {
        file.selected = !file.selected;
        renderList();
      });

      row.append(chk, label, del);
      fileListContainer.appendChild(row);
    });
  }

  document.getElementById('btnAddMod').addEventListener('click', () => {
    const list = [
      'C:\\Projects\\ExcelData\\연령별인구현황_A.xlsx',
      'C:\\Projects\\ExcelData\\연령별인구현황_B.xlsx'
    ];
    let added = 0;
    list.forEach(path => {
      if (!modFiles.some(f => f.path === path)) {
        modFiles.push({ path, checked: true, selected: false });
        added++;
      }
    });
    appendLog(added > 0 ? `수정본 파일 ${added}개가 추가되었습니다.` : '이미 목록에 있는 파일입니다.', added > 0 ? '#a6e3a1' : '#f9e2af');
    renderList();
  });

  document.getElementById('btnOpenMod').addEventListener('click', () => {
    const selected = modFiles.filter(f => f.selected);
    if (selected.length === 0) {
      alert('열어볼 수정본 파일을 목록에서 선택해 주세요.');
      return;
    }
    appendLog(`📖 수정본 엑셀 파일 열기: ${selected.map(f => fileName(f.path)).join(', ')}`, '#89b4fa');
  });

  document.getElementById('btnRemoveMod').addEventListener('click', () => {
    const removed = modFiles.filter(f => f.selected);
    if (removed.length === 0) {
      alert('삭제할 수정본 파일을 목록에서 선택해 주세요.');
      return;
    }
    modFiles = modFiles.filter(f => !f.selected);
    appendLog(`선택 삭제 완료: ${removed.map(f => fileName(f.path)).join(', ')}`, '#f9e2af');
    renderList();
  });

  document.getElementById('btnClearBoth').addEventListener('click', () => {
    basePath = '';
    txtBasePath.value = '';
    modFiles = [];
    renderList();
    appendLog('🧹 원본 파일 경로 및 수정본 파일 목록이 모두 비워졌습니다.', '#f9e2af');
  });

  // Run Aggregation Simulation
  document.getElementById('btnRun').addEventListener('click', () => {
    if (isBusy) return;
    if (!txtBasePath.value) {
      alert('원본 파일 경로를 선택해 주세요.');
      return;
    }

    let targets = modFiles.filter(f => f.checked);
    // Same rule as the desktop app: nothing checked but files listed → check and use all of them
    if (targets.length === 0 && modFiles.length > 0) {
      modFiles.forEach(f => { f.checked = true; });
      targets = modFiles.slice();
      renderList();
    }
    if (targets.length === 0) {
      alert('취합할 수정본 파일을 최소 1개 이상 목록에 추가해 주세요.');
      return;
    }

    let policy = Number(cboConflictPolicy.value);
    if (!POLICY_NAMES[policy]) {
      debugLog('WARN', `알 수 없는 충돌 정책 값(${cboConflictPolicy.value}) → 마지막 파일 우선으로 대체`);
      policy = ConflictPolicy.LastWins;
    }
    const policyName = cboConflictPolicy.options[cboConflictPolicy.selectedIndex].text;

    setBusy(true);
    terminalBody.replaceChildren();
    appendLog(`🚀 다중 파일 취합 작업 시작 (${targets.length}개 파일 선택됨)...`, '#a6e3a1');
    appendLog(`⚙️ 적용 충돌 정책: ${policyName}`, '#cba6f7');
    debugLog('DEBUG', `실행 설정: 원본=${txtBasePath.value} | 수정본=${targets.map(f => f.path).join(' ; ')} | 정책=${POLICY_NAMES[policy]}`);
    appendLog('🛡️ 원본 백업 사본 자동 생성 완료: 연령별인구현황_원본_backup.xlsx', '#a6e3a1');

    setTimeout(() => {
      try {
        // Demo conflict: two files touch the same cell → resolved by the selected policy
        const cellKey = 'DataSheet!R5C3';
        debugLog('DEBUG', `충돌 감지 [${cellKey}] 정책=${POLICY_NAMES[policy]} 기존='120' 수정='135'`);
        if (policy === ConflictPolicy.FirstWins) {
          appendLog(`    ℹ️ [충돌 스킵] [${cellKey}] 셀 기존 값(120) 보존 (수정본: 135)`);
        } else if (policy === ConflictPolicy.PromptUser) {
          const useNew = confirm(`셀 취합 충돌 발생\n\n시트 [DataSheet] R5C3\n기존 값: 120\n수정본 값: 135\n\n[확인] 수정본 값 적용 / [취소] 기존 값 유지`);
          debugLog('DEBUG', `충돌 처리 [${cellKey}] 사용자 선택=${useNew ? 'UseNewValue' : 'KeepExistingValue'}`);
          if (!useNew) appendLog(`    ℹ️ [사용자 선택] [${cellKey}] 셀 기존 값(120) 보존`);
        } else {
          debugLog('DEBUG', `충돌 처리 [${cellKey}] 마지막 파일 우선 → '135' 로 덮어씀`);
        }
        appendLog('  ├ 시트 [DataSheet]: 18개 셀 누적 반영 완료');
        appendLog('🎉 취합 완료! 모든 엑셀 취합 및 저장이 성공적으로 완료되었습니다.', '#a6e3a1');
      } catch (err) {
        appendLog(`❌ 오류 발생: ${err.message}`, '#f38ba8');
        debugLog('ERROR', '취합 작업 실패', err.stack);
      } finally {
        setBusy(false);
      }
    }, 800);
  });

  document.getElementById('btnReset').addEventListener('click', () => {
    if (isBusy) return;
    if (!txtBasePath.value) {
      alert('초기화할 원본 파일 경로를 먼저 선택해 주세요.');
      return;
    }
    if (!confirm(`선택한 원본 파일의 시트에서 취합으로 반영된 메모(주석)와 하이라이트 음영만 선택 초기화하시겠습니까?\n\n대상 파일: ${fileName(txtBasePath.value)}`)) {
      return;
    }
    setBusy(true);
    appendLog('🧹 원본 파일의 메모 및 음영 제거 작업 시작...', '#f9e2af');
    setTimeout(() => {
      appendLog('🎉 초기화 완료! 음영 및 메모가 깨끗하게 제거되었습니다.', '#a6e3a1');
      setBusy(false);
    }, 600);
  });
});
