<template>
  <div>
    <p data-text="ShawnLab">
      ShawnLab
    </p>
  </div>
</template>

<style scoped>
p {
  position: relative;
  margin: auto;
  font-size: 2rem;
  word-spacing: 0.2em;
  display: inline-block;
  line-height: 1;
  white-space: nowrap;
  color: transparent;
  background-color: #0052d9;
  background-clip: text;
}

p::after {
  content: attr(data-text);
  position: absolute;
  left: 0;
  top: 0;
  width: 100%;
  height: 100%;
  z-index: 5;
  background-image: linear-gradient(
    120deg,
    transparent 0%,
    transparent 1rem,
    #ef4723 6rem,
    transparent 6.15rem,
    transparent 10rem,
    rgba(255, 255, 255, 0.3) 15rem,
    transparent 20rem,
    transparent 22rem,
    rgba(255, 255, 255, 0.6) 27rem,
    white 28rem,
    rgba(255, 255, 255, 0.3) 28.15rem,
    transparent 33rem,
    transparent 35rem,
    rgba(255, 255, 255, 0.3) 40rem,
    transparent 45rem,
    transparent 100%
  );
  background-clip: text;
  background-size: 150% 100%;
  background-repeat: no-repeat;
  animation: shine 5s infinite linear;
}

@keyframes shine {
  0% {
    background-position: 50% 0;
  }
  100% {
    background-position: -190% 0;
  }
}
</style>
