// Deklaracje dla `npm run check` (TypeScript sprawdza JS, nic nie trafia do apki).
// shortcut: moduły dzielą się przez window.* i biblioteki z <script> bez typów, więc są `any`;
// gdy jakiś moduł dostanie typy, dopisz go tu konkretnie.
interface Window { [key: string]: any }
declare var XLSX: any;
declare var JSZip: any;
declare var LanguageModel: any;
