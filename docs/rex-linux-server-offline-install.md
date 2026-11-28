# 任务：在离线 Linux 服务器上安装 Rex（Superlogical）Linux 服务端

> 给执行 agent 的说明：你运行在一台 Windows 机器上，这台 Windows 能 SSH 到目标 Linux 服务器。目标 Linux **不能访问外网**，所以**不要**在 Linux 上执行任何 `curl ... | sh`、`apt`、`yum`、`dnf` 之类需要联网的命令。
> 每一步执行完都要检查输出。**某一步失败就停下来，向用户报告完整输出，不要自行改用其他方案。**
> 需要 `sudo` 的步骤先问用户同意再做。
>
> **全程离线**：你（agent）和目标 Linux 都不能访问外网，整个流程也不需要外网。文中出现的所有 URL 只是写给人看的来源说明，**不要尝试访问或下载**。只用到：Windows 本地文件 + 内网 SSH/scp。

## 开始前由用户提供的信息

| 项目 | 值（由用户填写） |
|---|---|
| 安装包在 Windows 上的完整路径 | `C:\...\rex-linux-amd64.tar.gz` |
| Linux 服务器 IP / 主机名 | `<LinuxIP>` |
| SSH 登录用户（也是以后用 Rex 连接的用户） | `<用户>` |
| 该用户是否可以使用 sudo | 是 / 否 |

如果用户没提供上面这些信息，先向用户询问，不要猜。安装包文件不存在时，**停止并报告**，不要尝试去下载。

---

## 0. 背景信息（已核实）

- **软件**：Rex，Superlogical 公司做的终端复用器（类似 tmux）。官网文档：https://www.superlogical.com/rex/docs/linux/installation
- **注意重名**：Debian/Ubuntu/Fedora 软件源里也有一个叫 `rex` 的 **Perl 自动化工具**，跟这个不是同一个软件，都会占用 `rex` 这个命令名。
- **Linux 版只能当服务端**：其他机器上的 Rex 客户端（目前主要是 macOS）通过 SSH 连进来使用。
- **官方状态**：“Rex on Linux is in testing”，属于测试阶段，不同版本之间的行为可能有变化。
- **安装包结构**（我已下载并检查过 x64 包）：
  - 文件名：`rex-linux-amd64.tar.gz`，约 17.9 MB
  - 解压后**只有一个文件**：`rex`（约 48 MB，Go 编译的 ELF 64 位可执行文件）
  - 动态依赖只有基础系统库：`libc.so.6`、`libresolv.so.2`、`libpthread.so.0`、`libdl.so.2`、`librt.so.1`，**不需要额外安装依赖**
  - 运行时会按需加载 `libpam.so.0`（只有系统级“登录服务器”模式需要，个人模式一般用不到）
  - 我检查的版本：preview 通道，build `202b1b3fc0219883112258a8d0f9f95e2d47daa9`
    该版本的原始下载地址（仅供人参考，agent 不要访问）：`https://blob.superlogical.com/download/rex/preview/202b1b3fc0219883112258a8d0f9f95e2d47daa9/rex-linux-amd64.tar.gz`
    SHA256：`41988e86e87e8b0b47c8bca7652d9959178360f7789785232090784e38e23cb5`
  - ARM64 包 `rex-linux-arm64.tar.gz` 我下载时返回 403，没能核实
- **官方系统要求**：
  - CPU：x86-64（必须支持 SSE4.2）或 ARM64
  - glibc ≥ 2.26（对应 RHEL 8、Ubuntu 18.04、Debian 10、Amazon Linux 2 及以上版本）
  - 开机自启需要 systemd；没有 systemd 也能运行 `rex` 命令

---

## 1. 预检查（在 Linux 上执行，只读，不改动任何东西）

先从 Windows 用 SSH 登录 Linux：`ssh <用户>@<LinuxIP>`，然后执行：

```bash
uname -m                                   # 必须是 x86_64（如果是 aarch64，就要换 arm64 的包）
grep -c sse4_2 /proc/cpuinfo               # 必须大于 0
ldd --version | head -1                    # glibc 版本必须 ≥ 2.26
cat /etc/os-release | head -4
ps -p 1 -o comm=                           # 输出 systemd 说明可以设开机自启
ssh -V                                     # 记下 OpenSSH 版本
command -v rex && rex -v 2>&1 | head -2    # 检查有没有重名的 rex
echo "PATH=$PATH"
df -h ~ /tmp | tail -2                     # 至少要有 100 MB 空间
```

**判断规则**：
- `uname -m` 不是 x86_64，或者 glibc < 2.26，或者 sse4_2 为 0：**停止**，向用户报告。
- `command -v rex` 有输出：如果 `rex -v` 打印出 `(R)?ex`，说明装的是 Perl 版 Rex。**停止**，问用户怎么处理：卸载 Perl 版，还是把新的 Rex 装到别的路径。

---

## 2. 在 Windows 上校验安装包

在 PowerShell 中执行（路径换成实际下载位置）：

```powershell
certutil -hashfile "C:\path\to\rex-linux-amd64.tar.gz" SHA256
```

- 如果下载的是上面那个 preview build，哈希必须是 `41988e86e87e8b0b47c8bca7652d9959178360f7789785232090784e38e23cb5`。
- 如果是别的版本，哈希会不一样。这种情况下把哈希值记下来，传到 Linux 后再对比一次，保证传输前后一致即可。

---

## 3. 把安装包传到 Linux

在 Windows PowerShell 中执行：

```powershell
scp "C:\path\to\rex-linux-amd64.tar.gz" <用户>@<LinuxIP>:/tmp/
```

如果没有 scp，可以用 WinSCP / MobaXterm 等工具通过 SFTP 上传到 `/tmp/`。

---

## 4. 解压并安装二进制（在 Linux 上执行）

```bash
sha256sum /tmp/rex-linux-amd64.tar.gz      # 和 Windows 上的结果对比，必须一致
mkdir -p ~/rex-install && tar -xzf /tmp/rex-linux-amd64.tar.gz -C ~/rex-install
ls -l ~/rex-install                        # 应该只有一个 rex 文件
file ~/rex-install/rex                     # 应该显示 ELF 64-bit ... x86-64
ldd ~/rex-install/rex | grep "not found"   # 应该没有任何输出
```

然后选一个安装位置：

**方案 A（推荐，需要 sudo，先问用户）**：装到 `/usr/local/bin`。这个目录在非交互式 SSH 的 PATH 里，客户端连接时才能找到 `rex`。
```bash
sudo install -m 0755 ~/rex-install/rex /usr/local/bin/rex
```

**方案 B（不需要 sudo）**：装到 `~/.local/bin`。
```bash
mkdir -p ~/.local/bin && install -m 0755 ~/rex-install/rex ~/.local/bin/rex
```
选方案 B 时要注意：非交互式 SSH 可能不加载 `~/.bashrc`，导致 PATH 里没有 `~/.local/bin`。第 7 步会检查这个问题。如果检查不通过，需要把 `export PATH="$HOME/.local/bin:$PATH"` 加到 `~/.bashrc` 的**最前面**，也就是 `case $- in *i*) ;; *) return;; esac` 这类“非交互就直接返回”的判断之前。

验证：
```bash
hash -r; command -v rex; rex version
```

---

## 5. 安装终端类型描述（terminfo）

```bash
rex terminfo setup
```

- 不加参数时只给当前用户安装。加 `--system` 是给全系统安装，需要 sudo，先问用户。
- 这个命令会根据 OpenSSH 版本决定要不要写 SSH 兼容配置：OpenSSH 低于 7.8 时会自动跳过。
- **禁止**自己手动在 Ubuntu 18.04 的 `/etc/ssh/ssh_config` 里加 `Include /etc/ssh/ssh_config.d/*`，官方文档说这会把所有账户的 ssh 弄坏。
- 执行完后检查一下 ssh 配置还能正常解析：`ssh -G example.com >/dev/null && echo ok`，必须输出 `ok`。

---

## 6. 安装并启动个人服务端（不需要 sudo）

**用以后要从客户端连进来的那个用户**来执行：

```bash
# SSH 会话里经常没有设置 XDG_RUNTIME_DIR，没有它 rex server install 会失败
[ -z "$XDG_RUNTIME_DIR" ] && export XDG_RUNTIME_DIR=/run/user/$(id -u)
ls -ld "$XDG_RUNTIME_DIR"                  # 目录必须存在

rex server install                         # 会安装 systemd 用户服务 rex.service，并设为开机启动
```

让服务在用户退出登录后继续运行：
```bash
loginctl enable-linger "$USER"             # 如果提示权限不足，就问用户是否同意执行 sudo loginctl enable-linger <用户>
loginctl show-user "$USER" -p Linger       # 应该输出 Linger=yes
```

**如果第 1 步发现系统没有 systemd**：跳过上面的命令，改为执行 `rex server start`。这种方式没有开机自启，服务器重启后，第一次有客户端连接时会再拉起服务。

验证：
```bash
rex server status
systemctl --user status rex --no-pager
journalctl --user -u rex -n 50 --no-pager
```

---

## 7. 从 Windows 端验证 SSH 调用

在 **Windows** 上执行：

```powershell
ssh <用户>@<LinuxIP> rex version
```

- 必须直接输出版本号，不能提示 `command not found`。如果找不到命令，回到第 4 步方案 B 的说明修复 PATH。
- Rex 客户端连接时**不能输入密码**，必须用 SSH 密钥登录。如果密钥有密码保护，要先 `ssh-add` 把密钥加到 ssh-agent 里。

客户端的连接写法（供用户参考）：
- 命令行：`rex -S ssh://<用户>@<LinuxIP> ls`
- macOS App：选择 “Add Remote Host…”，填 `ssh://<用户>@<LinuxIP>`

---

## 8. 安全要求（必须遵守）

- **不要**开启 HTTP 监听（TCP 7297 端口，`--allow-anonymous`）。这种方式是匿名访问，任何能连到这个端口的人都可以用你的身份执行程序。只用 SSH 方式。
- **不要**开启 Tailscale 等外网访问功能（离线环境本来也用不了）。
- 不要修改防火墙，不要动系统级的 `rex.service`、`rex-login-server.*` 等服务。这次只装个人服务端。

---

## 9. 升级、回滚、卸载

**升级**：用新的 tarball 重复第 2～4 步，覆盖原来的二进制文件，然后执行：
```bash
systemctl --user restart rex     # 没有 systemd 的话用：rex server stop（下次连接时会自动用新版本启动）
rex version
```

**卸载**：官方文档没有写卸载步骤，下面是根据安装内容推断出来的，执行前要先问用户：
```bash
systemctl --user disable --now rex
ls ~/.config/systemd/user/ | grep -i rex   # 先看清楚有哪些文件，再删掉 rex 相关的 unit 文件
systemctl --user daemon-reload
rm -f /usr/local/bin/rex ~/.local/bin/rex  # 删 /usr/local/bin 下的文件需要 sudo
# 数据和配置目录（删前要确认）：~/.local/state/rex  ~/.config/rex
```

---

## 10. 常见问题

| 现象 | 处理方法 |
|---|---|
| `rex server install` 报错，提到 XDG_RUNTIME_DIR 或 bus | 执行 `export XDG_RUNTIME_DIR=/run/user/$(id -u)` 后重试；如果这个目录不存在，先执行 `sudo loginctl enable-linger <用户>`（要先问用户） |
| `rex -v` 打印出 `(R)?ex` | 运行的是 Perl 版 Rex，`which -a rex` 查看所有同名命令，然后问用户怎么处理 |
| `cannot execute binary file` / `Exec format error` | CPU 架构和安装包不匹配 |
| 报 `GLIBC_x.xx not found` | glibc 版本太低，停止并报告 |
| 退出登录后服务停了 | 没开 linger，回到第 6 步 |
| 客户端连不上，但 ssh 本身正常 | 检查 `ssh <用户>@<IP> rex version`，常见原因是 PATH 问题，或者没有用密钥登录 |
| 看日志 | `journalctl --user -u rex`；`rex server status` 也会显示日志位置 |

---

## 完成后向用户汇报

1. 第 1 步的预检查输出
2. 二进制安装在哪个路径，`rex version` 的输出
3. `rex server status` 和 `systemctl --user status rex` 的输出
4. 从 Windows 执行 `ssh <用户>@<IP> rex version` 的结果
5. 哪些步骤跳过了，原因是什么

---

参考来源（仅供人查阅，agent 不要访问）：
- 官方安装文档：https://www.superlogical.com/rex/docs/linux/installation
- exe.dev 部署指南：https://exe.dev/deploy/rex
- NixOS 使用指南：https://frankper.gitlab.io/the-one-nix/guides/how-to-use-rex/
- 下载链接来源（LINUX DO 论坛）：https://linux.do/t/topic/3205583
