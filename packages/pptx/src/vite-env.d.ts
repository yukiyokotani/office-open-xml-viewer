/// <reference types="vite/client" />

// Vite's ?worker&inline query — not picked up automatically by tsc
declare module '*?worker&inline' {
  const WorkerFactory: { new (): Worker };
  export default WorkerFactory;
}
