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
  background: rgba(30, 41, 59, 0.8);
  padding: 2rem;
  border-radius: 1rem;
  width: 100%;
  max-width: 400px;
  text-align: center;
  border: 1px solid rgba(148, 163, 184, 0.2);
}

h2 {
  margin-bottom: 0.5rem;
  color: #e2e8f0;
}

.subtitle {
  color: #94a3b8;
  margin-bottom: 2rem;
  font-size: 0.9rem;
}

.login-input {
  width: 100%;
  padding: 0.75rem;
  border-radius: 0.5rem;
  border: 1px solid rgba(148, 163, 184, 0.4);
  background: rgba(15, 23, 42, 0.6);
  color: white;
  margin-bottom: 1rem;
  outline: none;
  transition: border-color 0.2s;
}

.login-input:focus {
  border-color: #6366f1;
}

.login-btn {
  width: 100%;
  padding: 0.75rem;
  background: #6366f1;
  color: white;
  border: none;
  border-radius: 0.5rem;
  cursor: pointer;
  font-weight: 600;
  transition: background 0.2s;
}

.login-btn:hover {
  background: #4f46e5;
}

.hint {
  font-size: 0.8rem;
  color: #64748b;
  margin-top: 0.5rem;
}

.error-msg {
  color: #ef4444;
  margin-bottom: 1rem;
  font-size: 0.9rem;
}
</style>
