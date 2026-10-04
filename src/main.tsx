import React from 'react';
import { createRoot } from 'react-dom/client';
import App from './App';
import './styles.css';

const root = document.getElementById('root');
if (!root) throw new Error('页面缺少应用容器');
createRoot(root).render(<React.StrictMode><App /></React.StrictMode>);
