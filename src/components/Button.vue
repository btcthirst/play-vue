<script setup lang="ts">

</script>

<template>
  <button class="border">
    <slot></slot>
  </button>
</template>

<style scoped>
  button {
  border-radius: 9px;
  border: 1px solid transparent;
  padding: 0.6em 1.2em;
  font-size: 1em;
  font-weight: 500;
  font-family: inherit;
  background-color: #2691f5;
  cursor: pointer;
  transition: border-color 0.25s;
}
button:hover {
  border-color: #555bdd;
}
button:focus,
button:focus-visible {
  outline: 4px auto -webkit-focus-ring-color;
}
</style>