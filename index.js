// Lean London compatibility entrypoint.
// Render's existing service may explicitly start `node index.js`.
// Keep this tiny wrapper so both `node index.js` and `npm start` launch the same lean backend.
import './lean-server.js';
