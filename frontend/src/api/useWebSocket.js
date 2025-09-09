// ...new file...
import { ref } from 'vue';

export function useWebSocket(wsUrlBase = 'ws://localhost:8000/ws/excel') {
  const ws = ref(null);
  const isConnected = ref(false);

  const connect = (instanceId) => {
    if (ws.value) ws.value.close();
    ws.value = new WebSocket(`${wsUrlBase}/${instanceId}`);
    ws.value.onopen = () => { isConnected.value = true; };
    ws.value.onclose = () => { isConnected.value = false; };
    ws.value.onerror = () => { isConnected.value = false; };
    return ws.value;
  };

  const sendMessage = (payload) => {
    if (ws.value && isConnected.value) {
      ws.value.send(JSON.stringify(payload));
    }
  };

  const close = () => {
    if (ws.value) ws.value.close();
    ws.value = null;
    isConnected.value = false;
  };

  return { connect, sendMessage, close, isConnected, ws };
}