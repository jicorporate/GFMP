// Test DOM Events & Interactivity in app.js
const fs = require('fs');
const path = require('path');

// Read index.html content
const htmlContent = fs.readFileSync(path.join(__dirname, '../mobile-web/index.html'), 'utf8');

// Minimal mock DOM parser / simulated environment
class MockClassList {
  constructor() { this.classes = new Set(); }
  add(c) { this.classes.add(c); }
  remove(c) { this.classes.delete(c); }
  toggle(c, force) {
    if (force === undefined) {
      if (this.classes.has(c)) this.classes.delete(c);
      else this.classes.add(c);
    } else if (force) this.classes.add(c);
    else this.classes.delete(c);
  }
  contains(c) { return this.classes.has(c); }
}

class MockElement {
  constructor(tag, id = '') {
    this.tagName = tag.toUpperCase();
    this.id = id;
    this.classList = new MockClassList();
    this.dataset = {};
    this.style = {};
    this.listeners = {};
    this.children = [];
    this._textContent = '';
    this._innerHTML = '';
    this.value = '';
    this.placeholder = '';
  }

  get textContent() { return this._textContent; }
  set textContent(v) { this._textContent = String(v); }

  get innerHTML() { return this._innerHTML; }
  set innerHTML(v) { this._innerHTML = String(v); }

  addEventListener(event, fn) {
    if (!this.listeners[event]) this.listeners[event] = [];
    this.listeners[event].push(fn);
  }

  removeEventListener(event, fn) {
    if (this.listeners[event]) {
      this.listeners[event] = this.listeners[event].filter(f => f !== fn);
    }
  }

  dispatchEvent(event) {
    const list = this.listeners[event.type] || [];
    list.forEach(fn => fn(event));
  }

  click() {
    this.dispatchEvent({ type: 'click', target: this });
  }

  scrollIntoView() {}
  appendChild(child) { this.children.push(child); }
  querySelector() { return null; }
  querySelectorAll() { return []; }
}

console.log('DOM mock ready.');
