<script setup>
import { useRouter } from 'vue-router'

const router = useRouter()

const goHome = () => {
  router.push('/')
}
</script>

<template>
  <div class="not-found-container">
    <div class="content-wrapper">
      <div class="error-code">404</div>
      <h1 class="title">页面未找到</h1>
      <p class="description">
        你似乎来到了知识荒原的边缘。<br />
        这里没有信号，也没有你想要的内容。
      </p>
      
      <div class="actions">
        <button @click="goHome" class="home-btn">
          <span class="icon">🚀</span>
          返回首页
        </button>
      </div>
    </div>
  </div>
</template>

<style scoped>
.not-found-container {
  min-height: 70vh;
  display: flex;
  align-items: center;
  justify-content: center;
  padding: 2rem;
  text-align: center;
  color: var(--color-body);
}

.content-wrapper {
  max-width: 600px;
  animation: fadeIn 0.8s ease-out;
}

.error-code {
  font-size: 8rem;
  font-weight: 900;
  line-height: 1;
  background: linear-gradient(135deg, #6366f1, #8b5cf6, #ec4899);
  -webkit-background-clip: text;
  background-clip: text;
  color: transparent;
  margin-bottom: 1rem;
  font-family: 'Segoe UI', sans-serif;
  text-shadow: 0 10px 30px rgba(99, 102, 241, 0.3);
}

.title {
  font-size: 2rem;
  margin-bottom: 1.5rem;
  color: var(--color-heading);
}

.description {
  font-size: 1.1rem;
  line-height: 1.6;
  color: var(--color-muted);
  margin-bottom: 3rem;
}

.home-btn {
  display: inline-flex;
  align-items: center;
  gap: 0.5rem;
  padding: 0.8rem 2rem;
  background: linear-gradient(135deg, #4f46e5, #7c3aed);
  color: white;
  border: none;
  border-radius: 99px;
  font-size: 1rem;
  font-weight: 600;
  cursor: pointer;
  transition: all 0.3s ease;
  box-shadow: 0 4px 15px rgba(79, 70, 229, 0.4);
  text-decoration: none;
}

.home-btn:hover {
  transform: translateY(-2px);
  box-shadow: 0 8px 25px rgba(79, 70, 229, 0.5);
  background: linear-gradient(135deg, #4338ca, #6d28d9);
}

.icon {
  font-size: 1.2rem;
}

@keyframes fadeIn {
  from {
    opacity: 0;
    transform: translateY(20px);
  }
  to {
    opacity: 1;
    transform: translateY(0);
  }
}
</style>
