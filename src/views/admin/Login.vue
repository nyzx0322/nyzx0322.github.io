<script setup>
import { ref } from 'vue'
import { useRouter } from 'vue-router'

const router = useRouter()
const password = ref('')
const errorMsg = ref('')

const handleLogin = async () => {
  errorMsg.value = ''
  try {
    const res = await fetch('/__api/check-auth', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ password: password.value })
    })
    
    if (!res.ok) {
      throw new Error('密码错误或服务不可用')
    }
    
    const data = await res.json()
    if (data.success) {
      window.localStorage.setItem('admin_token', data.token)
      router.push('/admin/')
    } else {
      throw new Error(data.error || '验证失败')
    }
  } catch (e) {
    errorMsg.value = e.message
  }
}
</script>

<template>
  <div class="login-container">
    <div class="login-card">
      <h2>后台管理登录</h2>
      <p class="subtitle">仅限本地开发环境使用</p>
      
      <form @submit.prevent="handleLogin">
        <div class="form-group">
          <input 
            v-model="password" 
            type="password" 
            placeholder="请输入管理员密码"
            class="login-input"
          >
        </div>
        
        <div v-if="errorMsg" class="error-msg">{{ errorMsg }}</div>
        
        <button type="submit" class="login-btn">
          进入系统
        </button>
      </form>
    </div>
  </div>
</template>

<style scoped>
.login-container {
  display: flex;
  justify-content: center;
  align-items: center;
  min-height: 80vh;
}

.login-card {
  background: var(--bg-login-card);
  padding: 2rem;
  border-radius: 1rem;
  width: 100%;
  max-width: 400px;
  text-align: center;
  border: 1px solid var(--border-login-card);
}

h2 {
  margin-bottom: 0.5rem;
  color: var(--color-heading);
}

.subtitle {
  color: var(--color-muted);
  margin-bottom: 2rem;
  font-size: 0.9rem;
}

.login-input {
  width: 100%;
  padding: 0.75rem;
  border-radius: 0.5rem;
  border: 1px solid var(--input-border);
  background: var(--input-bg);
  color: var(--input-text);
  margin-bottom: 1rem;
  outline: none;
  transition: border-color 0.2s;
}

.login-input:focus {
  border-color: var(--btn-primary-bg);
}

.login-btn {
  width: 100%;
  padding: 0.75rem;
  background: var(--btn-primary-bg);
  color: var(--btn-primary-text);
  border: none;
  border-radius: 0.5rem;
  cursor: pointer;
  font-weight: 600;
  transition: background 0.2s;
}

.login-btn:hover {
  background: var(--btn-primary-hover-bg);
}

.hint {
  font-size: 0.8rem;
  color: var(--color-muted);
  margin-top: 0.5rem;
}

.error-msg {
  color: var(--color-error);
  margin-bottom: 1rem;
  font-size: 0.9rem;
}
</style>
