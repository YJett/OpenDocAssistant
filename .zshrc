# 首先设置基本的 PATH，确保系统命令可用
export PATH="/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin:$PATH"

# Homebrew
eval "$(/opt/homebrew/bin/brew shellenv)"

# Oh My Zsh 配置
export ZSH="/Users/ywootae/.oh-my-zsh"
ZSH_THEME="robbyrussell"
plugins=(git)
source $ZSH/oh-my-zsh.sh

# iterm 插件
source /opt/homebrew/share/zsh-syntax-highlighting/zsh-syntax-highlighting.zsh
source /opt/homebrew/share/zsh-autosuggestions/zsh-autosuggestions.zsh

# autojump
[ -f /opt/homebrew/etc/profile.d/autojump.sh ] && . /opt/homebrew/etc/profile.d/autojump.sh

# Homebrew 镜像设置
export HOMEBREW_PIP_INDEX_URL=https://pypi.tuna.tsinghua.edu.cn/simple
export HOMEBREW_API_DOMAIN=https://mirrors.tuna.tsinghua.edu.cn/homebrew-bottles/api
export HOMEBREW_BOTTLE_DOMAIN=https://mirrors.ustc.edu.cn/homebrew-bottles

# C/C++
alias gcc='gcc-13'
alias g++='g++-13'
alias c++='c++-13'

# MySQL
export PATH="/opt/homebrew/opt/mysql@8.0/bin:$PATH"

# Java
export JAVA_HOME="/opt/homebrew/opt/openjdk@17"
export PATH="$JAVA_HOME/bin:$PATH"
export CLASSPATH="$JAVA_HOME/lib/tools.jar:$JAVA_HOME/lib/dt.jar:."

# Maven
export M2_HOME=/opt/homebrew/Cellar/maven/3.9.9
export PATH="$PATH:$M2_HOME/bin"

# Node.js
export NVM_DIR="$HOME/.nvm"
[ -s "/opt/homebrew/opt/nvm/nvm.sh" ] && \. "/opt/homebrew/opt/nvm/nvm.sh"
[ -s "/opt/homebrew/opt/nvm/etc/bash_completion.d/nvm" ] && \. "/opt/homebrew/opt/nvm/etc/bash_completion.d/nvm"
export PATH="/opt/homebrew/opt/node@22/bin:$PATH"

# Neo4j
export NEO4J_HOME=/Users/ywootae/Desktop/lab/neo4j
export PATH="$PATH:$NEO4J_HOME/bin"

# Node.js 选项
export NODE_OPTIONS="--openssl-legacy-provider"

# Conda 初始化
__conda_setup="$('/opt/homebrew/anaconda3/bin/conda' 'shell.zsh' 'hook' 2> /dev/null)"
if [ $? -eq 0 ]; then
    eval "$__conda_setup"
else
    if [ -f "/opt/homebrew/anaconda3/etc/profile.d/conda.sh" ]; then
        . "/opt/homebrew/anaconda3/etc/profile.d/conda.sh"
    else
        export PATH="/opt/homebrew/anaconda3/bin:$PATH"
    fi
fi
unset __conda_setup

# Redis 别名
alias redisstart='sudo launchctl start io.redis.redis-server'
alias redisstop='sudo launchctl stop io.redis.redis-server'

# 创建并使用环境变量文件存储敏感信息
if [ ! -f ~/.env_secrets ]; then
    touch ~/.env_secrets
    chmod 600 ~/.env_secrets
    cat > ~/.env_secrets << EOF
export DB_PASSWORD="your_password"
export ALIOSS_ACCESS_KEY_ID="your_key_id"
export ALIOSS_ACCESS_KEY_SECRET="your_key_secret"
export OPENAI_API_KEY="your_api_key"
EOF
fi

# 加载环境变量
source ~/.env_secrets 