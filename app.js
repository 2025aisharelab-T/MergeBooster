document.addEventListener('DOMContentLoaded', () => {
  let basePath = '';
  let modFiles = [];

  const txtBasePath = document.getElementById('txtBasePath');
  const fileListContainer = document.getElementById('fileListContainer');
  const emptyPlaceholder = document.getElementById('emptyPlaceholder');
  const fileCountBadge = document.getElementById('fileCountBadge');
  const terminalBody = document.getElementById('terminalBody');

  function appendLog(msg, color = '#89dceb') {
    const p = document.createElement('p');
    const time = new Date().toLocaleTimeString('ko-KR', { hour12: false });
    p.style.color = color;
    p.textContent = `[${time}] ${msg}`;
    terminalBody.appendChild(p);
    terminalBody.scrollTop = terminalBody.scrollHeight;
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

  // Render Modified Files List
  function renderList() {
    fileListContainer.innerHTML = '';
    if (modFiles.length === 0) {
      fileListContainer.appendChild(emptyPlaceholder);
      fileCountBadge.textContent = '0개 파일 선택됨';
      return;
    }
    fileCountBadge.textContent = `${modFiles.length}개 파일 선택됨`;

    modFiles.forEach((file, idx) => {
      const div = document.createElement('div');
      div.style.cssText = 'display: flex; align-items: center; justify-content: space-between; padding: 0.5rem 0.75rem; background-color: #f8fafc; border: 1px solid #e2e8f0; border-radius: 8px; font-size: 13px; font-weight: 500; color: #334155;';
      div.innerHTML = `
        <span style="overflow: hidden; text-overflow: ellipsis; white-space: nowrap;">📄 ${file}</span>
        <button class="btn-del-file" data-idx="${idx}" style="color: #ef4444; font-weight: 700; cursor: pointer; padding: 0.2rem 0.5rem; border-radius: 4px;">✕</button>
      `;
      fileListContainer.appendChild(div);
    });

    document.querySelectorAll('.btn-del-file').forEach(btn => {
      btn.addEventListener('click', (e) => {
        const i = parseInt(e.target.getAttribute('data-idx'));
        const removed = modFiles.splice(i, 1);
        appendLog(`목록에서 제거됨: ${removed[0]}`, '#f9e2af');
        renderList();
      });
    });
  }

  document.getElementById('btnAddMod').addEventListener('click', () => {
    const list = [
      'C:\\Projects\\ExcelData\\연령별인구현황_A.xlsx',
      'C:\\Projects\\ExcelData\\연령별인구현황_B.xlsx'
    ];
    list.forEach(f => {
      if (!modFiles.includes(f)) modFiles.push(f);
    });
    appendLog(`수정본 파일 ${list.length}개가 추가되었습니다.`, '#a6e3a1');
    renderList();
  });

  document.getElementById('btnOpenMod').addEventListener('click', () => {
    if (modFiles.length === 0) {
      alert('목록에서 수정본 파일을 추가해 주세요.');
      return;
    }
    appendLog(`📖 수정본 엑셀 파일 열기: ${modFiles.join(', ')}`, '#89b4fa');
  });

  document.getElementById('btnRemoveMod').addEventListener('click', () => {
    if (modFiles.length === 0) return;
    const removed = modFiles.pop();
    appendLog(`선택 삭제 완료: ${removed}`, '#f9e2af');
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
    if (!txtBasePath.value) {
      alert('원본 파일 경로를 먼저 지정해 주세요.');
      return;
    }
    if (modFiles.length === 0) {
      alert('취합할 수정본 파일을 최소 1개 이상 추가해 주세요.');
      return;
    }

    const cbo = document.getElementById('cboConflictPolicy');
    const policyName = cbo.options[cbo.selectedIndex].text;

    appendLog(`🚀 다중 파일 취합 실행 (${modFiles.length}개 파일)...`, '#a6e3a1');
    appendLog(`⚙️ 적용 충돌 처리 정책: ${policyName}`, '#cba6f7');
    appendLog(`🛡️ 원본 백업 사본 자동 생성 완료: 연령별인구현황_원본_backup.xlsx`, '#a6e3a1');

    setTimeout(() => {
      appendLog('  ├ 시트 [DataSheet]: 18개 셀 누적 반영 완료');
      appendLog('🎉 모든 엑셀 취합 및 저장이 성공적으로 완료되었습니다!', '#a6e3a1');
    }, 800);
  });

  document.getElementById('btnReset').addEventListener('click', () => {
    if (!txtBasePath.value) {
      alert('초기화할 원본 파일 경로를 선택해 주세요.');
      return;
    }
    appendLog('🧹 원본 파일의 메모 및 음영 제거 작업 실행...', '#f9e2af');
    setTimeout(() => {
      appendLog('🎉 초기화 완료! 음영 및 메모가 깨끗하게 제거되었습니다.', '#a6e3a1');
    }, 600);
  });
});
