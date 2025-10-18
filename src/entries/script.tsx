import React, { useEffect } from 'react';
import ReactDOM from 'react-dom/client';
import { initScript } from 'dingtalk-docs-cool-app';

function App() {
  useEffect(() => {
    // 初始化文档模型服务
    initScript({
      scriptUrl: new URL(`${process.env.PUBLIC_URL}/static/js/script.code.js`, window.location.href),
      onError: (e) => {
        console.log(e);
      }
    });
  }, []);
  return null;
}

const root = ReactDOM.createRoot(document.getElementById('root_script')!);
root.render(
  <React.StrictMode>
    <App/>
  </React.StrictMode>
);
