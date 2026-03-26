// App.js
import React, { useState } from 'react';
import './App.css';
import WordToPdfConverter from './components/WordToPdfConverter';

function App() {
  return (
    <div className="App">
      <WordToPdfConverter />
    </div>
  );
}

export default App;