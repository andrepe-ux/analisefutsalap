// app.js (Atualizado com a contagem de entradas e saídas por atleta)
let currentPeriod = 1;
let totalSeconds = 20 * 60;
let timerInterval = null;
let isRunning = false;

let homeGoals = 0;
let awayGoals = 0;

let actionLogsHistory = [];

let timeoutsUsed = {
    1: { home: false, away: false },
    2: { home: false, away: false }
};

let statsData = {
    home: { livres: 0, cantos: 0, lancamentos: 0, posse: 0, passes_falhados: 0, passes_completos: 0 },
    away: { livres: 0, cantos: 0, lancamentos: 0, posse: 0, passes_falhados: 0, passes_completos: 0 }
};

let players = [];
for (let i = 1; i <= 16; i++) {
    players.push({
        number: i,
        name: `Player ${i}`,
        isOnField: false,
        secondsPlayed: 0,
        secondsRested: 0,
        substitutionsCount: 0 // Contagem de vezes que entra em campo
    });
}

window.onload = function() {
    renderPlayersList();
    updateTimerDisplay();
    updateTimeoutUI();
};

function updateTeamNames() {
    let homeName = document.getElementById('input-home-name').value || 'CASA';
    let awayName = document.getElementById('input-away-name').value || 'VISITANTE';

    document.getElementById('home-team-title').innerText = `Plantel / Jogadores (${homeName.toUpperCase()})`;
    document.getElementById('stats-home-title').innerText = homeName.toUpperCase();
    document.getElementById('stats-away-title').innerText = awayName.toUpperCase();
    
    document.getElementById('to-home-label').innerText = `Time-out ${homeName}:`;
    document.getElementById('to-away-label').innerText = `Time-out ${awayName}:`;

    document.getElementById('pitch-remate-home-title').innerText = `REMATE ${homeName.toUpperCase()}`;
    document.getElementById('pitch-remate-away-title').innerText = `REMATE ${awayName.toUpperCase()}`;
    document.getElementById('pitch-golo-home-title').innerHTML = `GOLO ${homeName.toUpperCase()} (<span id="home-score">${homeGoals}</span>)`;
    document.getElementById('pitch-golo-away-title').innerHTML = `GOLO ${awayName.toUpperCase()} (<span id="away-score">${awayGoals}</span>)`;
}

// -------------------------------------------------------------
// GESTÃO DE SESSÃO EM JSON (Guardar e Importar)
// -------------------------------------------------------------
function exportSessionJSON() {
    let sessionData = {
        currentPeriod: currentPeriod,
        totalSeconds: totalSeconds,
        homeGoals: homeGoals,
        awayGoals: awayGoals,
        homeName: document.getElementById('input-home-name').value,
        awayName: document.getElementById('input-away-name').value,
        timeoutsUsed: timeoutsUsed,
        statsData: statsData,
        players: players,
        actionLogsHistory: actionLogsHistory,
        htmlLogs: document.getElementById('action-log').innerHTML,
        pitchMarkers: {
            homeShot: document.getElementById('pitch-shot-home').innerHTML,
            homeGoal: document.getElementById('pitch-goal-home').innerHTML,
            awayShot: document.getElementById('pitch-shot-away').innerHTML,
            awayGoal: document.getElementById('pitch-goal-away').innerHTML
        }
    };

    let dataStr = "data:text/json;charset=utf-8," + encodeURIComponent(JSON.stringify(sessionData, null, 2));
    let downloadAnchor = document.createElement('a');
    downloadAnchor.setAttribute("href", dataStr);
    downloadAnchor.setAttribute("download", `sessao_futsal_${sessionData.homeName}_vs_${sessionData.awayName}.json`);
    document.body.appendChild(downloadAnchor);
    downloadAnchor.click();
    downloadAnchor.remove();
}

function importSessionJSON(event) {
    let file = event.target.files[0];
    if (!file) return;

    let reader = new FileReader();
    reader.onload = function(e) {
        try {
            let sessionData = JSON.parse(e.target.result);

            pauseTimer();
            currentPeriod = sessionData.currentPeriod || 1;
            totalSeconds = sessionData.totalSeconds || (20 * 60);
            homeGoals = sessionData.homeGoals || 0;
            awayGoals = sessionData.awayGoals || 0;

            document.getElementById('input-home-name').value = sessionData.homeName || 'CASA';
            document.getElementById('input-away-name').value = sessionData.awayName || 'VISITANTE';
            
            timeoutsUsed = sessionData.timeoutsUsed || { 1: { home: false, away: false }, 2: { home: false, away: false } };
            statsData = sessionData.statsData || statsData;
            players = sessionData.players || players;
            actionLogsHistory = sessionData.actionLogsHistory || [];

            document.getElementById('home-score').innerText = homeGoals;
            document.getElementById('away-score').innerText = awayGoals;
            document.getElementById('action-log').innerHTML = sessionData.htmlLogs || '';

            if (sessionData.pitchMarkers) {
                document.getElementById('pitch-shot-home').innerHTML = sessionData.pitchMarkers.homeShot;
                document.getElementById('pitch-goal-home').innerHTML = sessionData.pitchMarkers.homeGoal;
                document.getElementById('pitch-shot-away').innerHTML = sessionData.pitchMarkers.awayShot;
                document.getElementById('pitch-goal-away').innerHTML = sessionData.pitchMarkers.awayGoal;
            }

            document.getElementById('btn-p1').className = currentPeriod === 1 ? 'period-btn active' : 'period-btn';
            document.getElementById('btn-p2').className = currentPeriod === 2 ? 'period-btn active' : 'period-btn';

            for (let t of ['home', 'away']) {
                for (let k in statsData[t]) {
                    let el = document.getElementById(`${t}-${k}`);
                    if (el) el.innerText = statsData[t][k];
                }
            }

            updateTeamNames();
            updateTimerDisplay();
            updateTimeoutUI();
            renderPlayersList();

            alert("Sessão recuperada com sucesso!");
        } catch (err) {
            alert("Erro ao ler o ficheiro JSON. Certifique-se de que é um ficheiro de sessão válido.");
            console.error(err);
        }
    };
    reader.readAsText(file);
}

// -------------------------------------------------------------
// MODELO E IMPORTAÇÃO EXCEL
// -------------------------------------------------------------
function downloadTemplate() {
    let templateData = [
        ["Configuração da Partida", ""],
        ["Equipa Casa", "CASA"],
        ["Equipa Visitante", "VISITANTE"],
        [],
        ["Número (Coluna A)", "Nome do Atleta (Coluna B)"]
    ];

    for (let i = 1; i <= 16; i++) {
        templateData.push([i, `Atleta ${i}`]);
    }

    let ws = XLSX.utils.aoa_to_sheet(templateData);
    let wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Plantel e Equipas");
    XLSX.writeFile(wb, "modelo_futsal_completo.xlsx");
}

function importExcel(event) {
    let file = event.target.files[0];
    if (!file) return;

    let reader = new FileReader();
    reader.onload = function(e) {
        let data = new Uint8Array(e.target.result);
        let workbook = XLSX.read(data, { type: 'array' });
        let firstSheetName = workbook.SheetNames[0];
        let worksheet = workbook.Sheets[firstSheetName];
        let json = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

        let importedPlayers = 0;

        for (let i = 0; i < json.length; i++) {
            let colA = json[i][0];
            let colB = json[i][1];

            if (colA && String(colA).toLowerCase().includes("equipa casa") && colB) {
                document.getElementById('input-home-name').value = colB;
            }
            if (colA && String(colA).toLowerCase().includes("equipa visitante") && colB) {
                document.getElementById('input-away-name').value = colB;
            }

            if (colA !== undefined && colB !== undefined && !isNaN(colA) && importedPlayers < 16) {
                players[importedPlayers].number = colA;
                players[importedPlayers].name = String(colB).trim();
                importedPlayers++;
            }
        }

        updateTeamNames();
        renderPlayersList();
        alert(`Importação concluída com sucesso! ${importedPlayers} atletas carregados.`);
    };
    reader.readAsArrayBuffer(file);
}

// -------------------------------------------------------------
// CRONÓMETRO E PERÍODOS
// -------------------------------------------------------------
function switchPeriod(period) {
    pauseTimer();
    currentPeriod = period;
    totalSeconds = 20 * 60;
    
    document.getElementById('btn-p1').className = period === 1 ? 'period-btn active' : 'period-btn';
    document.getElementById('btn-p2').className = period === 2 ? 'period-btn active' : 'period-btn';
    
    updateTimerDisplay();
    updateTimeoutUI();
    logAction('SISTEMA', `Início do ${period}º Período`, null, null);
}

function updateTimerDisplay() {
    let minutes = Math.floor(totalSeconds / 60);
    let seconds = totalSeconds % 60;
    document.getElementById('timer').innerText = 
        `${String(minutes).padStart(2, '0')}:${String(seconds).padStart(2, '0')}`;
}

function startTimer() {
    if (!isRunning && totalSeconds > 0) {
        isRunning = true;
        timerInterval = setInterval(() => {
            if (totalSeconds > 0) {
                totalSeconds--;
                updateTimerDisplay();

                players.forEach(player => {
                    if (player.isOnField) {
                        player.secondsPlayed++;
                    } else {
                        player.secondsRested++;
                    }
                });

                updateTimesOnly();
            } else {
                pauseTimer();
                alert(`Fim do ${currentPeriod}º Período!`);
            }
        }, 1000);
    }
}

function pauseTimer() {
    isRunning = false;
    clearInterval(timerInterval);
}

function resetTimer() {
    pauseTimer();
    totalSeconds = 20 * 60;
    updateTimerDisplay();
}

function requestTimeout(team) {
    if (!timeoutsUsed[currentPeriod][team]) {
        timeoutsUsed[currentPeriod][team] = true;
        updateTimeoutUI();
        pauseTimer();
        let teamName = team === 'home' ? (document.getElementById('input-home-name').value || 'CASA') : (document.getElementById('input-away-name').value || 'VISITANTE');
        logAction(teamName.toUpperCase(), `Time-out pedido no ${currentPeriod}º Período`, null, null);
    } else {
        alert("Esta equipa já utilizou a pausa técnica neste período!");
    }
}

function updateTimeoutUI() {
    let homeUsed = timeoutsUsed[currentPeriod].home;
    document.getElementById('timeout-home-status').innerText = homeUsed ? 'Utilizado' : 'Disponível';
    document.getElementById('timeout-home-status').className = homeUsed ? 'to-used used' : 'to-used';
    document.getElementById('timeout-home-btn').disabled = homeUsed;

    let awayUsed = timeoutsUsed[currentPeriod].away;
    document.getElementById('timeout-away-status').innerText = awayUsed ? 'Utilizado' : 'Disponível';
    document.getElementById('timeout-away-status').className = awayUsed ? 'to-used used' : 'to-used';
    document.getElementById('timeout-away-btn').disabled = awayUsed;
}

// -------------------------------------------------------------
// REGISTOS NOS CAMPOS E ESTATÍSTICAS
// -------------------------------------------------------------
function handlePitchDoubleClick(event, team, type) {
    const pitch = event.currentTarget;
    const rect = pitch.getBoundingClientRect();
    const x = ((event.clientX - rect.left) / rect.width) * 100;
    const y = ((event.clientY - rect.top) / rect.height) * 100;

    addMarker(pitch, x, y, type);

    let teamName = team === 'home' ? (document.getElementById('input-home-name').value || 'CASA') : (document.getElementById('input-away-name').value || 'VISITANTE');

    if (type === 'goal') {
        if (team === 'home') {
            homeGoals++;
            document.getElementById('home-score').innerText = homeGoals;
        } else {
            awayGoals++;
            document.getElementById('away-score').innerText = awayGoals;
        }
        logAction(teamName.toUpperCase(), `GOLO (${currentPeriod}ºP)`, x.toFixed(0), y.toFixed(0));
    } else {
        logAction(teamName.toUpperCase(), `Remate (${currentPeriod}ºP)`, x.toFixed(0), y.toFixed(0));
    }
}

function addMarker(pitchElement, xPercent, yPercent, type) {
    const marker = document.createElement('div');
    marker.className = `pitch-marker ${type === 'goal' ? 'marker-goal' : 'marker-shot'}`;
    marker.style.left = `${xPercent}%`;
    marker.style.top = `${yPercent}%`;
    pitchElement.appendChild(marker);
}

function recordStat(team, actionType) {
    statsData[team][actionType]++;
    document.getElementById(`${team}-${actionType}`).innerText = statsData[team][actionType];
    
    let teamName = team === 'home' ? (document.getElementById('input-home-name').value || 'CASA') : (document.getElementById('input-away-name').value || 'VISITANTE');
    
    let actionNames = {
        livres: 'Falta/Livre',
        cantos: 'Canto',
        lancamentos: 'Lançamento',
        posse: 'Perda de Posse',
        passes_falhados: 'Passe Falhado',
        passes_completos: 'Passe Certo'
    };
    
    logAction(teamName.toUpperCase(), actionNames[actionType], null, null);
}

function logAction(teamName, actionDesc, coordX, coordY) {
    let currentClock = document.getElementById('timer').innerText;
    let logBox = document.getElementById('action-log');
    
    let fullDesc = actionDesc;
    if (coordX !== null && coordY !== null) {
        fullDesc += ` [X: ${coordX}%, Y: ${coordY}%]`;
    }

    actionLogsHistory.push({
        periodo: `${currentPeriod}ºP`,
        tempo: currentClock,
        equipa: teamName,
        acao: fullDesc,
        x: coordX,
        y: coordY
    });
    
    let entry = document.createElement('div');
    entry.className = 'log-entry';
    entry.innerHTML = `[${currentPeriod}ºP - ${currentClock}] <b>${teamName}</b>: ${fullDesc}`;
    
    logBox.insertBefore(entry, logBox.firstChild);
}

// -------------------------------------------------------------
// RELATÓRIOS (EXCEL E PDF)
// -------------------------------------------------------------
function exportReportExcel() {
    let homeName = document.getElementById('input-home-name').value || 'CASA';
    let awayName = document.getElementById('input-away-name').value || 'VISITANTE';

    let wb = XLSX.utils.book_new();

    let resumoData = [
        ["Relatório da Partida - Análise Futsal AP"],
        ["Equipa Casa", homeName, "Golos", homeGoals],
        ["Equipa Visitante", awayName, "Golos", awayGoals],
        [],
        ["Estatísticas", homeName, awayName],
        ["Faltas/Livres", statsData.home.livres, statsData.away.livres],
        ["Cantos", statsData.home.cantos, statsData.away.cantos],
        ["Lançamentos", statsData.home.lancamentos, statsData.away.lancamentos],
        ["Perdas de Posse", statsData.home.posse, statsData.away.posse],
        ["Passes Falhados", statsData.home.passes_falhados, statsData.away.passes_falhados],
        ["Passes Certos", statsData.home.passes_completos, statsData.away.passes_completos]
    ];
    let wsResumo = XLSX.utils.aoa_to_sheet(resumoData);
    XLSX.utils.book_append_sheet(wb, wsResumo, "Resumo");

    let acoesData = [["Período", "Tempo", "Equipa", "Ação / Evento", "Coord X (%)", "Coord Y (%)"]];
    actionLogsHistory.forEach(log => {
        acoesData.push([log.periodo, log.tempo, log.equipa, log.acao, log.x !== null ? log.x : '-', log.y !== null ? log.y : '-']);
    });
    let wsAcoes = XLSX.utils.aoa_to_sheet(acoesData);
    XLSX.utils.book_append_sheet(wb, wsAcoes, "Eventos e Campos");

    let playersData = [["Número", "Nome do Jogador", "Tempo em Jogo", "Tempo no Banco", "Nº de Entradas"]];
    players.forEach(p => {
        playersData.push([p.number, p.name, formatTime(p.secondsPlayed), formatTime(p.secondsRested), p.substitutionsCount]);
    });
    let wsPlayers = XLSX.utils.aoa_to_sheet(playersData);
    XLSX.utils.book_append_sheet(wb, wsPlayers, "Plantel");

    XLSX.writeFile(wb, `Relatorio_Futsal_${homeName}_vs_${awayName}.xlsx`);
}

async function exportReportPDF() {
    const { jsPDF } = window.jspdf;
    let doc = new jsPDF();

    let homeName = document.getElementById('input-home-name').value || 'CASA';
    let awayName = document.getElementById('input-away-name').value || 'VISITANTE';

    doc.setFont("helvetica", "bold");
    doc.setFontSize(16);
    doc.text("Relatorio Oficial - Análise Futsal AP", 14, 20);

    doc.setFontSize(12);
    doc.text(`Partida: ${homeName} ${homeGoals} - ${awayGoals} ${awayName}`, 14, 30);

    doc.setFont("helvetica", "normal");
    doc.setFontSize(10);
    doc.text(`Gerado em: ${new Date().toLocaleDateString()} por andrepe @ 2026`, 14, 38);

    let y = 48;
    doc.setFont("helvetica", "bold");
    doc.text("Estatísticas Coletivas:", 14, y);
    y += 6;
    doc.setFont("helvetica", "normal");
    doc.text(`- Faltas/Livres: ${homeName} (${statsData.home.livres}) x (${statsData.away.livres}) ${awayName}`, 14, y); y += 6;
    doc.text(`- Cantos: ${homeName} (${statsData.home.cantos}) x (${statsData.away.cantos}) ${awayName}`, 14, y); y += 6;
    doc.text(`- Perdas de Posse: ${homeName} (${statsData.home.posse}) x (${statsData.away.posse}) ${awayName}`, 14, y); y += 10;

    // SECÇÃO DOS TEMPOS E ENTRADAS DOS JOGADORES NO PDF
    doc.setFont("helvetica", "bold");
    doc.text("Estatísticas dos Jogadores (Tempos e Entradas):", 14, y);
    y += 6;
    doc.setFont("helvetica", "normal");

    players.forEach(p => {
        if (y > 280) {
            doc.addPage();
            y = 20;
        }
        doc.text(`Nº ${p.number} - ${p.name}: Jogo [${formatTime(p.secondsPlayed)}] | Banco [${formatTime(p.secondsRested)}] | Entradas [${p.substitutionsCount}]`, 14, y);
        y += 6;
    });

    y += 4;

    // CAPTURA DOS CAMPOS DE FUTSAL PARA O PDF
    doc.setFont("helvetica", "bold");
    doc.text("Mapas Visuais dos Campos (Remates e Golos):", 14, y);
    y += 6;

    try {
        let pitchContainer = document.querySelector('.pitch-container-grid');
        let canvas = await html2canvas(pitchContainer, { backgroundColor: '#1e1e1e', scale: 2 });
        let imgData = canvas.toDataURL('image/png');
        
        let imgWidth = 180;
        let imgHeight = (canvas.height * imgWidth) / canvas.width;
        
        if (y + imgHeight > 280) {
            doc.addPage();
            y = 20;
        }

        doc.addImage(imgData, 'PNG', 14, y, imgWidth, imgHeight);
        y += imgHeight + 10;
    } catch (err) {
        console.error("Erro ao capturar os campos para PDF", err);
    }

    if (y > 270) {
        doc.addPage();
        y = 20;
    }

    doc.setFont("helvetica", "bold");
    doc.text("Histórico de Ações Detalhado:", 14, y);
    y += 6;
    doc.setFont("helvetica", "normal");

    actionLogsHistory.forEach(log => {
        if (y > 280) {
            doc.addPage();
            y = 20;
        }
        doc.text(`[${log.periodo} - ${log.tempo}] ${log.equipa}: ${log.acao}`, 14, y);
        y += 6;
    });

    doc.save(`Relatorio_Futsal_${homeName}_vs_${awayName}.pdf`);
}

// -------------------------------------------------------------
// GESTÃO DO PLANTEL
// -------------------------------------------------------------
function togglePlayerField(index) {
    let player = players[index];
    if (player) {
        player.isOnField = !player.isOnField;
        // Incrementa sempre que o atleta entra em campo (passa de falso para verdadeiro)
        if (player.isOnField) {
            player.substitutionsCount++;
        }
        renderPlayersList();
    }
}

function updatePlayerNumber(index, newNum) {
    players[index].number = newNum;
}

function updatePlayerName(index, newName) {
    players[index].name = newName;
}

function formatTime(totalSecs) {
    let mins = Math.floor(totalSecs / 60);
    let secs = totalSecs % 60;
    return `${String(mins).padStart(2, '0')}:${String(secs).padStart(2, '0')}`;
}

function renderPlayersList() {
    const container = document.getElementById('players-container');
    container.innerHTML = '';

    players.forEach((player, index) => {
        let card = document.createElement('div');
        card.className = `player-card ${player.isOnField ? 'field' : ''}`;
        
        card.innerHTML = `
            <input type="text" class="player-num-input" value="${player.number}" onchange="updatePlayerNumber(${index}, this.value)">
            <input type="text" class="player-name-input" value="${player.name}" onchange="updatePlayerName(${index}, this.value)">
            <div class="player-times-box">
                <div class="time-item">
                    <span class="time-label">Jogo</span>
                    <span class="time-val play" id="play-time-${index}">${formatTime(player.secondsPlayed)}</span>
                </div>
                <div class="time-item">
                    <span class="time-label">Banco</span>
                    <span class="time-val rest" id="rest-time-${index}">${formatTime(player.secondsRested)}</span>
                </div>
            </div>
            <button class="btn-card" onclick="togglePlayerField(${index})" style="background-color: ${player.isOnField ? '#b30000' : '#0073e6'};" title="Entradas: ${player.substitutionsCount}">
                ${player.isOnField ? 'Sair' : 'Entrar'} (${player.substitutionsCount})
            </button>
        `;
        
        container.appendChild(card);
    });
}

function updateTimesOnly() {
    players.forEach((player, index) => {
        const playElem = document.getElementById(`play-time-${index}`);
        const restElem = document.getElementById(`rest-time-${index}`);
        
        if (playElem) playElem.innerText = formatTime(player.secondsPlayed);
        if (restElem) restElem.innerText = formatTime(player.secondsRested);
    });
}
}
