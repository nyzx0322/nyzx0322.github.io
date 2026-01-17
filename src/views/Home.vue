<script setup>
import HelloWorld from '../components/HelloWorld.vue'
import { useCounter } from '../composables/useCounter'
import { useRootStore } from '../stores'

const { count, increment, decrement } = useCounter()
const { state } = useRootStore()
</script>

<template>
  <div class="home">
    <HelloWorld :msg="`首页 - ${state.siteName} v${state.version}`" />

    <section class="home-section">
      <h2>计数器示例（useCounter）</h2>
      <p>当前计数：{{ count }}</p>
      <button type="button" @click="increment">增加</button>
      <button type="button" @click="decrement">减少</button>
    </section>

    <section class="home-section">
      <h2>自动聚焦输入框（vFocus）</h2>
      <input v-focus type="text" placeholder="页面加载后自动获得焦点" />
    </section>
  </div>
</template>

<style scoped>
.home {
  display: flex;
  flex-direction: column;
  gap: 24px;
  align-items: center;
}

.home-section {
  border: 1px solid #e5e5e5;
  border-radius: 8px;
  padding: 16px 24px;
  min-width: 320px;
}

.home-section h2 {
  margin-top: 0;
}

.home-section button + button {
  margin-left: 8px;
}

.home-section input {
  width: 100%;
  max-width: 320px;
  padding: 8px 12px;
  border-radius: 4px;
  border: 1px solid #ccc;
  box-sizing: border-box;
}
</style>
