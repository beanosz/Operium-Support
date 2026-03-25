import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, simpledialog
import subprocess
import ctypes
import sys
import threading
import socket
import concurrent.futures
import os
import time

# --- VARIÁVEIS GLOBAIS ---
dispositivos_ad_cache = [] 

# --- VALIDAÇÃO DE ADMINISTRADOR ---
def is_admin():
    try: return ctypes.windll.shell32.IsUserAnAdmin()
    except: return False

# --- FUNÇÕES AUXILIARES ---
def log_msg(msg):
    log_area.insert(tk.END, f"{msg}\n")
    log_area.see(tk.END)

def limpar_log():
    log_area.delete(1.0, tk.END)
    log_msg("✔ Tela de logs limpa.")

def run_ps_command(cmd, timeout_sec=8):
    full_cmd = f'powershell -NoProfile -Command "{cmd}"'
    try:
        startupinfo = subprocess.STARTUPINFO()
        startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        result = subprocess.run(full_cmd, capture_output=True, text=True, shell=True, startupinfo=startupinfo, timeout=timeout_sec)
        if result.returncode == 0 and result.stdout.strip():
            return [line.strip() for line in result.stdout.strip().split('\n') if line.strip()]
        return []
    except Exception:
        return []

# --- INTEGRAÇÃO COM ACTIVE DIRECTORY (AGORA AUTOMÁTICO) ---
def varredura_ad():
    janela.after(0, lambda: log_msg("\n>>> CONSULTANDO ACTIVE DIRECTORY (Auto-Load)..."))
    janela.after(0, lambda: search_entry.config(state='disabled'))
    
    for item in tree_rede.get_children(): tree_rede.delete(item)
    dispositivos_ad_cache.clear()

    def rotina():
        try:
            cmd = "$searcher = [adsisearcher]'(objectCategory=computer)'; $searcher.PageSize = 2000; $searcher.FindAll() | ForEach-Object { $_.Properties.name[0] }"
            result = run_ps_command(cmd, timeout_sec=30)

            if not result:
                janela.after(0, lambda: log_msg("✖ Nenhum computador encontrado no AD ou falha de conexão."))
                return

            computadores = [c.upper() for c in result if c.strip()]
            janela.after(0, lambda: log_msg(f"⏳ {len(computadores)} máquinas listadas no AD. Resolvendo DNS e checando status online..."))

            def resolver_dns(hostname):
                try:
                    ip = socket.gethostbyname(hostname)
                    return (ip, hostname)
                except socket.gaierror:
                    return None 

            with concurrent.futures.ThreadPoolExecutor(max_workers=150) as executor:
                resultados = executor.map(resolver_dns, computadores)

            for res in resultados:
                if res: dispositivos_ad_cache.append(res)

            dispositivos_ad_cache.sort(key=lambda x: x[1])

            for ip, hostname in dispositivos_ad_cache:
                janela.after(0, tree_rede.insert, "", "end", values=(ip, hostname))

            janela.after(0, lambda: log_msg(f"✔ AD escaneado! {len(dispositivos_ad_cache)} máquinas online prontas para uso."))
        except Exception as e:
            janela.after(0, lambda: log_msg(f"✖ Erro de AD: {str(e)}"))
        finally:
            janela.after(0, lambda: search_entry.config(state='normal'))

    threading.Thread(target=rotina, daemon=True).start()

def filtrar_lista(*args):
    termo = search_var.get().lower()
    for item in tree_rede.get_children():
        tree_rede.delete(item)
    for ip, hostname in dispositivos_ad_cache:
        if termo in ip.lower() or termo in hostname.lower():
            tree_rede.insert("", "end", values=(ip, hostname))

# --- RASTREADOR DE USUÁRIOS ---
def localizar_usuario():
    alvo = search_user_var.get().strip().lower()
    if not alvo: 
        return messagebox.showwarning("Aviso", "Digite o nome (ou parte do nome) do usuário!")
    if not dispositivos_ad_cache: 
        return messagebox.showwarning("Aviso", "Aguarde o carregamento do Domínio primeiro!")

    log_msg(f"\n>>> 🕵️ RASTREANDO O USUÁRIO '{alvo}' NA REDE...")
    btn_buscar_user.config(state='disabled', text="⏳ RASTREANDO...")
    for item in tree_users.get_children(): tree_users.delete(item)

    def rotina_busca():
        encontrados = []
        
        def checar_pc(pc_info):
            ip, hostname = pc_info
            cmd = f"Get-WmiObject Win32_ComputerSystem -ComputerName '{ip}' | Select-Object -ExpandProperty UserName"
            res = run_ps_command(cmd, timeout_sec=4)
            if res and res[0] and alvo in res[0].lower():
                return (res[0], ip, hostname)
            return None

        with concurrent.futures.ThreadPoolExecutor(max_workers=100) as executor:
            resultados = executor.map(checar_pc, dispositivos_ad_cache)

        for r in resultados:
            if r: encontrados.append(r)

        if encontrados:
            for user_completo, ip, host in encontrados:
                janela.after(0, lambda u=user_completo, i=ip, h=host: tree_users.insert("", "end", values=(u, f"{h} ({i})")))
                janela.after(0, lambda u=user_completo, h=host: log_msg(f"🎯 ALVO ENCONTRADO: '{u}' está usando a máquina {h}"))
        else:
            janela.after(0, lambda: log_msg(f"✖ O usuário '{alvo}' não foi encontrado em nenhuma máquina ligada no momento."))

        janela.after(0, lambda: btn_buscar_user.config(state='normal', text="🎯 INICIAR CAÇADA"))

    threading.Thread(target=rotina_busca, daemon=True).start()

# --- LÓGICA DE INVENTÁRIO (POP-UP) ---
def abrir_popup_inventario(alvo, hostname_alvo=""):
    if not alvo: 
        return messagebox.showwarning("Aviso", "Selecione um computador primeiro!")
    if not hostname_alvo: hostname_alvo = alvo 

    log_msg(f">>> ABRINDO INVENTÁRIO DE: {hostname_alvo}")
    
    pop = tk.Toplevel(janela)
    pop.title(f"Inventário - {hostname_alvo}")
    pop.geometry("680x500")
    pop.configure(bg="#1e1e1e")
    pop.attributes("-topmost", True)

    tk.Label(pop, text=f"📦 Relatório de Sistema: {hostname_alvo}", font=("Segoe UI", 14, "bold"), bg="#1e1e1e", fg="white").pack(pady=10)
    
    txt_info = scrolledtext.ScrolledText(pop, bg="#0d0d0d", fg="#00ff99", font=("Consolas", 11), relief="flat")
    txt_info.pack(fill="both", expand=True, padx=15, pady=(0, 15))
    txt_info.insert(tk.END, f"⏳ Solicitando dados via WMI a {alvo}...\nIsso pode levar alguns segundos.\n")

    def rotina_coleta():
        try:
            ip_alvo = socket.gethostbyname(alvo)
            user_info = run_ps_command(f"Get-WmiObject Win32_ComputerSystem -ComputerName '{ip_alvo}' | Select-Object -ExpandProperty UserName")
            net_info = run_ps_command(f"Get-WmiObject Win32_NetworkAdapterConfiguration -ComputerName '{ip_alvo}' -Filter 'IPEnabled=True' | ForEach-Object {{ $_.IPAddress[0] + ' | ' + $_.MACAddress }}")
            os_info = run_ps_command(f"Get-WmiObject Win32_OperatingSystem -ComputerName '{ip_alvo}' | ForEach-Object {{ $_.Caption + ' (Build ' + $_.BuildNumber + ')' }}")
            cpu_info = run_ps_command(f"Get-WmiObject Win32_Processor -ComputerName '{ip_alvo}' | Select-Object -ExpandProperty Name")
            ram_info = run_ps_command(f"Get-WmiObject Win32_ComputerSystem -ComputerName '{ip_alvo}' | ForEach-Object {{ [math]::round($_.TotalPhysicalMemory / 1GB, 1) }}")
            disk_info = run_ps_command(f"Get-WmiObject Win32_LogicalDisk -ComputerName '{ip_alvo}' -Filter 'DriveType=3' | ForEach-Object {{ $_.DeviceID + ' Livre: ' + [math]::round($_.FreeSpace/1GB, 1) + 'GB / Total: ' + [math]::round($_.Size/1GB, 1) + 'GB' }}")
            bios_info = run_ps_command(f"Get-WmiObject Win32_Bios -ComputerName '{ip_alvo}' | ForEach-Object {{ $_.SerialNumber }}")
            phys_disk = run_ps_command(f"Get-CimInstance -Namespace Root\\Microsoft\\Windows\\Storage -ClassName MSFT_PhysicalDisk -ComputerName '{ip_alvo}' | ForEach-Object {{ if($_.MediaType -eq 4){{'SSD'}} elseif($_.MediaType -eq 3){{'HDD'}} else{{'Desconhecido'}} }}", timeout_sec=10)

            saida = f"\n{'-'*50}\n"
            if user_info and user_info[0]: saida += f"👤 USUÁRIO:   {user_info[0]}\n"
            else: saida += f"👤 USUÁRIO:   (Nenhum / Tela de Login)\n"

            if net_info:
                for net in net_info:
                    try: ip_placa, mac_placa = net.split(' | '); saida += f"🌐 REDE:      IP: {ip_placa}  MAC: {mac_placa}\n"
                    except: pass
            else: saida += f"🌐 REDE:      IP: {ip_alvo} (MAC Indisponível)\n"

            if bios_info: saida += f"🏷️ SERIAL:    {bios_info[0]}\n"

            saida += f"{'-'*50}\n"
            if os_info: saida += f"🖥️  OS:        {os_info[0]}\n"
            if cpu_info: saida += f"🧠 CPU:       {cpu_info[0]}\n"
            if ram_info: saida += f"💾 RAM:       {ram_info[0]} GB\n"
            if phys_disk:
                tipos = " + ".join(sorted(list(set(phys_disk)), reverse=True))
                saida += f"⚡ DISCO FÍSICO: [{tipos}]\n"
            if disk_info:
                for d in disk_info: saida += f"💿 DISCO LÓG.: {d}\n"
            if not os_info: saida += "✖ Falha: RPC Bloqueado ou Acesso Negado pelo Firewall.\n"
            
            pop.after(0, lambda: [txt_info.delete(1.0, tk.END), txt_info.insert(tk.END, saida)])
        except Exception as e:
            pop.after(0, lambda: txt_info.insert(tk.END, f"\n✖ Erro Crítico: {str(e)}"))

    threading.Thread(target=rotina_coleta).start()

# --- GERENCIADOR DE TAREFAS NATIVO ---
def abrir_monitor_desempenho():
    alvo = ip_entry.get().strip()
    if not alvo: return messagebox.showwarning("Aviso", "Selecione um computador primeiro!")
    try: ip_alvo = socket.gethostbyname(alvo)
    except socket.gaierror: return log_msg(f"✖ Falha na resolução DNS de {alvo}.")

    log_msg(f"\n>>> ABRINDO NATIVE TASK MANAGER: {alvo}")
    monitor = tk.Toplevel(janela); monitor.title(f"Gerenciador de Tarefas - {alvo}")
    monitor.geometry("850x650"); monitor.configure(bg="#202020")
    
    f_sidebar = tk.Frame(monitor, bg="#202020", width=200); f_sidebar.pack(side="left", fill="y"); f_sidebar.pack_propagate(False)
    tk.Label(f_sidebar, text="Desempenho", font=("Segoe UI", 14, "bold"), bg="#202020", fg="white", anchor="w").pack(fill="x", padx=15, pady=15)
    btn_cpu = tk.Frame(f_sidebar, bg="#333333", height=60); btn_cpu.pack(fill="x", padx=10, pady=5); btn_cpu.pack_propagate(False)
    tk.Label(btn_cpu, text="CPU", font=("Segoe UI", 11, "bold"), bg="#333333", fg="white", anchor="w").pack(fill="x", padx=10, pady=(10,0))
    lbl_side_cpu = tk.Label(btn_cpu, text="---%", font=("Segoe UI", 9), bg="#333333", fg="#aaaaaa", anchor="w"); lbl_side_cpu.pack(fill="x", padx=10)

    f_main = tk.Frame(monitor, bg="#111111"); f_main.pack(side="right", fill="both", expand=True)
    f_header = tk.Frame(f_main, bg="#111111"); f_header.pack(fill="x", padx=20, pady=20)
    f_title = tk.Frame(f_header, bg="#111111"); f_title.pack(side="left")
    tk.Label(f_title, text="CPU", font=("Segoe UI", 24), bg="#111111", fg="white").pack(anchor="w")
    tk.Label(f_title, text="% de Utilização", font=("Segoe UI", 9), bg="#111111", fg="#aaaaaa").pack(anchor="w")

    f_header_dir = tk.Frame(f_header, bg="#111111"); f_header_dir.pack(side="right", anchor="n")
    lbl_cpu_name = tk.Label(f_header_dir, text="Carregando hardware...", font=("Segoe UI", 12), bg="#111111", fg="white"); lbl_cpu_name.pack(anchor="e")
    lbl_status = tk.Label(f_header_dir, text="● Iniciando coleta...", font=("Segoe UI", 9), bg="#111111", fg="#aaaaaa"); lbl_status.pack(anchor="e")

    f_graph = tk.Frame(f_main, bg="#111111"); f_graph.pack(fill="x", padx=20)
    tk.Label(f_graph, text="100%", font=("Segoe UI", 8), bg="#111111", fg="#aaaaaa").pack(anchor="e")
    w_canvas, h_canvas = 580, 200
    canvas = tk.Canvas(f_graph, width=w_canvas, height=h_canvas, bg="#111111", highlightthickness=1, highlightbackground="#333333"); canvas.pack(pady=2)
    tk.Label(f_graph, text="60 segundos", font=("Segoe UI", 8), bg="#111111", fg="#aaaaaa").pack(anchor="w")

    f_stats = tk.Frame(f_main, bg="#111111"); f_stats.pack(fill="x", padx=20, pady=10)
    f_col1 = tk.Frame(f_stats, bg="#111111", width=120); f_col1.pack(side="left", fill="y", padx=(0, 20))
    tk.Label(f_col1, text="Utilização", font=("Segoe UI", 10), bg="#111111", fg="#aaaaaa", anchor="w").pack(fill="x")
    lbl_util = tk.Label(f_col1, text="---", font=("Segoe UI", 16), bg="#111111", fg="white", anchor="w"); lbl_util.pack(fill="x", pady=(0, 10))
    tk.Label(f_col1, text="Processos", font=("Segoe UI", 10), bg="#111111", fg="#aaaaaa", anchor="w").pack(fill="x")
    lbl_procs = tk.Label(f_col1, text="---", font=("Segoe UI", 14), bg="#111111", fg="white", anchor="w"); lbl_procs.pack(fill="x", pady=(0, 10))
    tk.Label(f_col1, text="Tempo de atividade", font=("Segoe UI", 10), bg="#111111", fg="#aaaaaa", anchor="w").pack(fill="x")
    lbl_up = tk.Label(f_col1, text="0:00:00:00", font=("Segoe UI", 12), bg="#111111", fg="white", anchor="w"); lbl_up.pack(fill="x", pady=(0, 5))

    f_col2 = tk.Frame(f_stats, bg="#111111", width=120); f_col2.pack(side="left", fill="y", padx=20)
    tk.Label(f_col2, text="Velocidade", font=("Segoe UI", 10), bg="#111111", fg="#aaaaaa", anchor="w").pack(fill="x")
    lbl_speed = tk.Label(f_col2, text="--- GHz", font=("Segoe UI", 16), bg="#111111", fg="white", anchor="w"); lbl_speed.pack(fill="x", pady=(0, 10))
    tk.Label(f_col2, text="Memória RAM", font=("Segoe UI", 10), bg="#111111", fg="#aaaaaa", anchor="w").pack(fill="x")
    lbl_ram = tk.Label(f_col2, text="---", font=("Segoe UI", 14), bg="#111111", fg="white", anchor="w"); lbl_ram.pack(fill="x")

    f_col3 = tk.Frame(f_stats, bg="#111111"); f_col3.pack(side="right", fill="y")
    lbl_base = tk.Label(f_col3, text="Velocidade base: --- GHz", font=("Segoe UI", 10), bg="#111111", fg="#aaaaaa", anchor="e"); lbl_base.pack(fill="x", pady=2)
    lbl_cores = tk.Label(f_col3, text="Núcleos: ---", font=("Segoe UI", 10), bg="#111111", fg="#aaaaaa", anchor="e"); lbl_cores.pack(fill="x", pady=2)
    lbl_logicos = tk.Label(f_col3, text="Processadores lógicos: ---", font=("Segoe UI", 10), bg="#111111", fg="#aaaaaa", anchor="e"); lbl_logicos.pack(fill="x", pady=2)

    cpu_hist = [0] * 60; monitor_ativo = [True]; buscando = [False]; hardware_carregado = [False]

    def fechar_monitor(): monitor_ativo[0] = False; monitor.destroy()
    monitor.protocol("WM_DELETE_WINDOW", fechar_monitor)

    def desenhar_grafico_nativo():
        canvas.delete("all")
        for i in range(1, 10): canvas.create_line(i * (w_canvas / 10), 0, i * (w_canvas / 10), h_canvas, fill="#2a2a2a")
        for i in range(1, 4): canvas.create_line(0, i * (h_canvas / 4), w_canvas, i * (h_canvas / 4), fill="#2a2a2a")
        passo_x = w_canvas / (len(cpu_hist) - 1); pontos_linha = []; pontos_poly = [(0, h_canvas)]
        for i, val in enumerate(cpu_hist):
            x = i * passo_x; y = max(1, min(h_canvas-1, h_canvas - (val / 100.0 * h_canvas)))
            pontos_linha.extend([x, y]); pontos_poly.extend([x, y])
        pontos_poly.extend([w_canvas, h_canvas])
        canvas.create_polygon(pontos_poly, fill="#1a3b5c", outline="")
        canvas.create_line(pontos_linha, fill="#429ce3", width=2, smooth=False)

    def aplicar_dados(resultado):
        if not monitor_ativo[0]: return
        if resultado and len(resultado) > 0:
            try:
                if not hardware_carregado[0] and len(resultado) > 4:
                    lbl_cpu_name.config(text=resultado[1]); lbl_cores.config(text=f"Núcleos: {resultado[2]}"); lbl_logicos.config(text=f"Processadores lógicos: {resultado[3]}")
                    try: lbl_base.config(text=f"Velocidade base: {(float(resultado[4]) / 1000.0):.2f} GHz")
                    except: pass
                    hardware_carregado[0] = True
                partes = resultado[0].split('|')
                c_val = float(partes[0].replace(',', '.')); r_val = float(partes[1].replace(',', '.')); procs = partes[2]; up_str = partes[3]
                try: curr_speed = float(partes[4].replace(',', '.')) / 1000.0
                except: curr_speed = 0.0

                lbl_util.config(text=f"{int(c_val)}%"); lbl_side_cpu.config(text=f"{int(c_val)}%")
                lbl_ram.config(text=f"{r_val:.1f}%"); lbl_procs.config(text=procs); lbl_up.config(text=up_str); lbl_speed.config(text=f"{curr_speed:.2f} GHz")
                cpu_hist.pop(0); cpu_hist.append(c_val); desenhar_grafico_nativo()
                lbl_status.config(text="○ Aguardando rede...", fg="#aaaaaa")
            except Exception: pass 
        buscando[0] = False
        if monitor_ativo[0]: monitor.after(1000, disparar_busca)

    def thread_busca():
        ps_script = (
            f"$cpu = Get-WmiObject Win32_Processor -ComputerName '{ip_alvo}'; "
            f"$os = Get-WmiObject Win32_OperatingSystem -ComputerName '{ip_alvo}'; "
            "if ($cpu -and $os) { "
            "  $load = $cpu | Measure-Object -Property LoadPercentage -Average | Select-Object -ExpandProperty Average; "
            "  $curr = $cpu | Measure-Object -Property CurrentClockSpeed -Average | Select-Object -ExpandProperty Average; "
            "  if ($null -eq $load) { $load = 0 }; if ($null -eq $curr) { $curr = 0 }; "
            "  $tot = [math]::Round([double]$os.TotalVisibleMemorySize/1048576, 1); $free = [math]::Round([double]$os.FreePhysicalMemory/1048576, 1); "
            "  $usoRAM = 0; if ($tot -gt 0) { $usoRAM = [math]::Round((($tot - $free) / $tot) * 100, 1) }; "
            "  $procs = $os.NumberOfProcesses; $boot = $os.ConvertToDateTime($os.LastBootUpTime); $now = $os.ConvertToDateTime($os.LocalDateTime); $up = $now - $boot; "
            "  $upStr = '{0}:{1:D2}:{2:D2}:{3:D2}' -f $up.Days, $up.Hours, $up.Minutes, $up.Seconds; "
            "  Write-Output ($load.ToString() + '|' + $usoRAM.ToString() + '|' + $procs.ToString() + '|' + $upStr + '|' + $curr.ToString()); "
        )
        if not hardware_carregado[0]: ps_script += "  $c = @($cpu)[0]; Write-Output $c.Name; Write-Output $c.NumberOfCores; Write-Output $c.NumberOfLogicalProcessors; Write-Output $c.MaxClockSpeed; "
        ps_script += "}"
        resultado = run_ps_command(ps_script)
        if monitor_ativo[0]: monitor.after(0, aplicar_dados, resultado)

    def disparar_busca():
        if monitor_ativo[0] and not buscando[0]:
            buscando[0] = True; lbl_status.config(text="● Atualizando dados...", fg="#00ff99")
            threading.Thread(target=thread_busca).start()

    desenhar_grafico_nativo(); disparar_busca()

# --- AÇÕES DE CONTROLE REMOTO ---
def gpupdate_remoto():
    alvo = ip_entry.get().strip()
    if not alvo: return messagebox.showwarning("Aviso", "Selecione um computador!")
    log_msg(f"\n>>> FORÇANDO ATUALIZAÇÃO DE POLÍTICA (GPUPDATE) EM: {alvo}")
    def rotina():
        cmd = f"Invoke-WmiMethod -Class Win32_Process -Name Create -ArgumentList 'cmd.exe /c gpupdate /force' -ComputerName '{alvo}'"
        run_ps_command(cmd, timeout_sec=15)
        janela.after(0, lambda: log_msg(f"✔ Gpupdate disparado com sucesso em {alvo}."))
    threading.Thread(target=rotina).start()

def acesso_remoto():
    alvo = ip_entry.get().strip()
    if not alvo: return messagebox.showwarning("Aviso", "Selecione um computador!")
    log_msg(f"\n>>> DISPARANDO ACESSO REMOTO PARA: {alvo}")
    try: os.system(f"start msra.exe /offerra {alvo}"); log_msg("✔ Solicitação enviada ao Windows.")
    except Exception as e: log_msg(f"✖ Erro: {str(e)}")

# --- REBOOT E STRESS TEST ---
def reiniciar_maquina():
    alvo = ip_entry.get().strip()
    if not alvo: return messagebox.showwarning("Aviso", "Selecione um computador!")
    if not messagebox.askyesno("Reiniciar PC", f"A máquina {alvo} será reiniciada IMEDIATAMENTE. Continuar?"): return
    log_msg(f"\n>>> REINICIANDO: {alvo}")
    try:
        if subprocess.run(f"shutdown /r /f /t 0 /m \\\\{alvo}", shell=True, capture_output=True).returncode == 0: log_msg("✔ Comando de reinício aceito!")
        else: log_msg("✖ Falha ao reiniciar (Firewall ou Offline).")
    except Exception as e: log_msg(f"✖ Erro: {str(e)}")

def reinicio_em_loop():
    alvo = ip_entry.get().strip()
    if not alvo: return messagebox.showwarning("Aviso", "Selecione um computador!")

    ciclos = simpledialog.askinteger("Stress Test", f"Quantas vezes deseja reiniciar {alvo} em sequência?\n(Máximo recomendado: 5)", minvalue=1, maxvalue=5)
    if not ciclos: return

    if not messagebox.askyesno("⚠️ ATENÇÃO - MODO STRESS TEST", f"Você vai forçar a reinicialização de {alvo} por {ciclos} vezes seguidas.\n\nQualquer trabalho não salvo pelo usuário será perdido. Tem certeza absoluta que deseja iniciar este ciclo?"):
        return

    log_msg(f"\n>>> [STRESS TEST] INICIANDO CICLO DE REBOOT ({ciclos}x) EM: {alvo}")

    def rotina_loop():
        for i in range(ciclos):
            janela.after(0, lambda msg=f"\n🔄 [Ciclo {i+1}/{ciclos}] Disparando reinício...": log_msg(msg))
            try: subprocess.run(f"shutdown /r /f /t 0 /m \\\\{alvo}", shell=True, capture_output=True)
            except Exception as e:
                janela.after(0, lambda msg=f"✖ Erro ao enviar comando: {str(e)}": log_msg(msg))
                break

            janela.after(0, lambda: log_msg("⏳ Passo 1/3: Aguardando a máquina cair (Ping falhar)..."))
            
            desligou = False
            for _ in range(40): 
                resp = subprocess.run(f"ping -n 1 -w 1000 {alvo}", shell=True, capture_output=True)
                if resp.returncode != 0: 
                    desligou = True
                    break
                time.sleep(2)
                
            if desligou: janela.after(0, lambda: log_msg("✔ Máquina offline. Passo 2/3: Aguardando Boot..."))
            else: janela.after(0, lambda: log_msg("⚠️ A máquina demorou para desligar, monitorando..."))

            wmi_pronto = False
            tempo_espera = 0
            janela.after(0, lambda: log_msg("⏳ Passo 3/3: Sondando serviços RPC/WMI (Windows ligando)..."))
            
            while tempo_espera < 300: 
                cmd_teste = f"Get-WmiObject Win32_OperatingSystem -ComputerName '{alvo}' | Select-Object -ExpandProperty Caption"
                teste = run_ps_command(cmd_teste, timeout_sec=3)
                if teste and len(teste) > 0:
                    wmi_pronto = True
                    janela.after(0, lambda: log_msg("🟢 SISTEMA 100% CARREGADO! WMI Respondendo."))
                    break
                time.sleep(4); tempo_espera += 4

            if not wmi_pronto:
                janela.after(0, lambda: log_msg(f"✖ TIMEOUT CRÍTICO: {alvo} travou no boot ou Firewall bloqueou. Abortando."))
                break
            
            if i < ciclos - 1:
                janela.after(0, lambda: log_msg("⚡ Preparando próximo tiro na sequência..."))
                time.sleep(2)

        janela.after(0, lambda: log_msg(f"\n🏁 FIM DO STRESS TEST. A máquina {alvo} completou os ciclos com sucesso."))

    threading.Thread(target=rotina_loop, daemon=True).start()

def reset_spooler():
    log_msg("\n>>> RESET SPOOLER (LOCAL)...")
    subprocess.run("net stop spooler && net start spooler", shell=True, creationflags=subprocess.CREATE_NO_WINDOW)
    log_msg("✔ Serviço reiniciado.")

def renovar_ip():
    log_msg("\n>>> RENOVAR IP (LOCAL)...")
    subprocess.run("ipconfig /release && ipconfig /renew", shell=True, creationflags=subprocess.CREATE_NO_WINDOW)
    log_msg("✔ IP Atualizado.")

# --- INTERAÇÕES DA LISTA ---
def ao_clique_lista(event):
    item = tree_rede.selection()
    if item:
        ip = tree_rede.item(item[0], "values")[0]
        ip_entry.delete(0, tk.END)
        ip_entry.insert(0, ip)

def ao_duplo_clique_lista(event):
    item = tree_rede.selection()
    if item:
        ip, hostname = tree_rede.item(item[0], "values")
        ip_entry.delete(0, tk.END)
        ip_entry.insert(0, ip)
        abrir_popup_inventario(ip, hostname)

def ao_duplo_clique_lista_users(event):
    item = tree_users.selection()
    if item:
        val_maquina = tree_users.item(item[0], "values")[1] 
        ip_extraido = val_maquina.split("(")[1].replace(")", "")
        host_extraido = val_maquina.split(" ")[0]
        ip_entry.delete(0, tk.END)
        ip_entry.insert(0, ip_extraido)
        abrir_popup_inventario(ip_extraido, host_extraido)

# --- INICIALIZAÇÃO DA INTERFACE ---
def iniciar_interface():
    global log_area, ip_entry, janela, tree_rede, search_var, search_entry, search_user_var, tree_users, btn_buscar_user
    janela = tk.Tk()
    janela.title("Operium Suport v5.0 - The Automation Update")
    janela.geometry("1200x720")
    janela.configure(bg="#2d2d2d")

    style = ttk.Style()
    style.theme_use("clam")
    style.configure("Treeview", background="#1e1e1e", foreground="white", fieldbackground="#1e1e1e", borderwidth=0)
    style.map("Treeview", background=[('selected', '#007acc')])
    
    # --- CORREÇÃO DO DESIGN DAS ABAS ---
    style.configure("TNotebook", background="#1e1e1e", borderwidth=0)
    
    style.configure("TNotebook.Tab", 
                    background="#333333",  # Cinza mais escuro para a aba inativa (fundo)
                    foreground="#888888",  # Texto meio apagado para tirar o foco
                    font=("Segoe UI", 10, "bold"), 
                    padding=[15, 4],       # Padding base
                    borderwidth=0)         # Tira a borda feia do Windows antigo

    style.map("TNotebook.Tab", 
              background=[('selected', '#007acc')], # Fica azulão quando clica
              foreground=[('selected', '#ffffff')], # O texto acende pra branco
              expand=[('selected', [0, 4, 0, 0])]   # O PULO DO GATO: Empurra a aba 4 pixels para CIMA
             )

    frame_esq = tk.Frame(janela, bg="#1e1e1e", width=350)
    frame_esq.pack(side="left", fill="y", padx=10, pady=10)

    # --- CRIANDO AS ABAS ANTES DE ADICIONAR CONTEÚDO ---
    abas = ttk.Notebook(frame_esq)
    
    aba_pcs = tk.Frame(abas, bg="#1e1e1e")
    aba_users = tk.Frame(abas, bg="#1e1e1e")

    # --- ABA COMPUTADORES ---
    frame_busca = tk.Frame(aba_pcs, bg="#1e1e1e")
    frame_busca.pack(fill="x", pady=(10, 5))
    tk.Label(frame_busca, text="🔍 Filtrar:", bg="#1e1e1e", fg="white").pack(side="left")
    
    search_var = tk.StringVar()
    search_var.trace_add("write", filtrar_lista)
    search_entry = tk.Entry(frame_busca, textvariable=search_var, bg="#404040", fg="white", insertbackground="white")
    search_entry.pack(side="right", fill="x", expand=True, padx=5)

    tree_rede = ttk.Treeview(aba_pcs, columns=("IP", "Hostname"), show="headings")
    tree_rede.heading("IP", text="Endereço IP"); tree_rede.heading("Hostname", text="Nome")
    tree_rede.column("IP", width=110); tree_rede.column("Hostname", width=160)
    tree_rede.pack(fill="both", expand=True, pady=5)
    
    tree_rede.bind("<ButtonRelease-1>", ao_clique_lista)
    tree_rede.bind("<Double-1>", ao_duplo_clique_lista)

    # --- ABA USUÁRIOS ---
    tk.Label(aba_users, text="Digite o login (ex: joao.silva):", bg="#1e1e1e", fg="white").pack(anchor="w", pady=(10, 2))
    search_user_var = tk.StringVar()
    tk.Entry(aba_users, textvariable=search_user_var, bg="#404040", fg="white", font=("Consolas", 11), insertbackground="white").pack(fill="x", pady=5)
    
    btn_buscar_user = tk.Button(aba_users, text="🎯 INICIAR CAÇADA", bg="#e83e8c", fg="white", font=("Segoe UI", 9, "bold"), relief="flat", command=localizar_usuario)
    btn_buscar_user.pack(fill="x", pady=5)

    tree_users = ttk.Treeview(aba_users, columns=("User", "IP"), show="headings")
    tree_users.heading("User", text="Usuário"); tree_users.heading("IP", text="Máquina/IP")
    tree_users.column("User", width=120); tree_users.column("IP", width=150)
    tree_users.pack(fill="both", expand=True, pady=5)

    tree_users.bind("<Double-1>", ao_duplo_clique_lista_users)

    # --- ADICIONANDO AS ABAS CONSTRUÍDAS AO NOTEBOOK ---
    abas.add(aba_pcs, text="💻 Computadores")
    abas.add(aba_users, text="👤 Caçar Usuário")
    abas.pack(fill="both", expand=True)

    # --- LADO DIREITO (COMANDOS) ---
    frame_dir = tk.Frame(janela, bg="#2d2d2d")
    frame_dir.pack(side="right", fill="both", expand=True, padx=10, pady=10)
    tk.Label(frame_dir, text="OPERIUM SUPORT", font=("Segoe UI", 26, "bold"), bg="#2d2d2d", fg="#ffffff").pack(anchor="w")
    
    frame_acoes = tk.LabelFrame(frame_dir, text=" Central de Comandos ", bg="#2d2d2d", fg="#aaaaaa", padx=10, pady=10)
    frame_acoes.pack(fill="x", pady=10)
    
    frame_input = tk.Frame(frame_acoes, bg="#2d2d2d"); frame_input.pack(fill="x", pady=(0, 10))
    tk.Label(frame_input, text="Alvo:", bg="#2d2d2d", fg="white", font=("Segoe UI", 12)).pack(side="left")
    ip_entry = tk.Entry(frame_input, width=25, font=("Consolas", 14), bg="#1e1e1e", fg="#00ff99"); ip_entry.pack(side="left", padx=10)
    
    tk.Button(frame_input, text="🧹 LIMPAR TELA", bg="#4d4d4d", fg="white", font=("Segoe UI", 9, "bold"), relief="flat", command=limpar_log).pack(side="right", padx=10)

    frame_botoes1 = tk.Frame(frame_acoes, bg="#2d2d2d"); frame_botoes1.pack(fill="x", pady=2)
    frame_botoes2 = tk.Frame(frame_acoes, bg="#2d2d2d"); frame_botoes2.pack(fill="x", pady=2)

    btn_style = {"font": ("Segoe UI", 9, "bold"), "width": 16, "pady": 5, "relief": "flat"}
    
    tk.Button(frame_botoes1, text="🔍 INVENTÁRIO", bg="#007acc", fg="white", command=lambda: abrir_popup_inventario(ip_entry.get()), **btn_style).pack(side="left", padx=3)
    tk.Button(frame_botoes1, text="📊 DESEMPENHO", bg="#17a2b8", fg="white", command=abrir_monitor_desempenho, **btn_style).pack(side="left", padx=3)
    tk.Button(frame_botoes1, text="🖥️ ACESSO REMOTO", bg="#6f42c1", fg="white", command=acesso_remoto, **btn_style).pack(side="left", padx=3)
    tk.Button(frame_botoes1, text="🔄 GPUPDATE", bg="#e83e8c", fg="white", command=gpupdate_remoto, **btn_style).pack(side="left", padx=3)

    tk.Button(frame_botoes2, text="⚡ REINICIAR PC", bg="#ff8c00", fg="white", command=reiniciar_maquina, **btn_style).pack(side="left", padx=3)
    tk.Button(frame_botoes2, text="⚠️ STRESS TEST", bg="#dc3545", fg="white", command=reinicio_em_loop, **btn_style).pack(side="left", padx=3)
    tk.Button(frame_botoes2, text="♻️ SPOOLER LOCAL", bg="#d94e36", fg="white", command=reset_spooler, **btn_style).pack(side="left", padx=3)
    tk.Button(frame_botoes2, text="🌐 RENOVAR IP", bg="#28a745", fg="white", command=renovar_ip, **btn_style).pack(side="left", padx=3)

    log_area = scrolledtext.ScrolledText(frame_dir, bg="#000000", fg="#00ff00", font=("Consolas", 10), relief="flat")
    log_area.pack(fill="both", expand=True, pady=10)
    
    log_msg("Sistema Operium Online v5.0. Interface carregada.")
    
    # Executa a varredura AD automaticamente após 500ms (tempo para a tela desenhar e não travar)
    janela.after(500, varredura_ad)
    
    janela.mainloop()

if __name__ == "__main__":
    if is_admin(): iniciar_interface()
    else: ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, " ".join(sys.argv), None, 1)