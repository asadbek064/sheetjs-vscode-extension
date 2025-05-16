import * as assert from 'assert';
import * as vscode from 'vscode';
import * as path from 'path';
import * as fs from 'fs';
import * as os from 'os';
import * as XLSX from 'xlsx';
import { LRUCache } from '../cacheManagement/lruCache';
import { WorkbookCache } from '../cacheManagement/workbookCache';


// test suite for the LRUCache class
suite('LRUCache Tests', () => {
	test('Should store and retrieve values', () => {
		const cache = new LRUCache<string, number>(3);
		cache.set('a', 1);
		cache.set('b', 2);

		assert.strictEqual(cache.get('a'), 1);
		assert.strictEqual(cache.get('b'), 2);
		assert.strictEqual(cache.get('c'), undefined);
	});

	test('Should respect maximum size', () => {
		const cache = new LRUCache<string, number>(2);
		cache.set('a', 1);
		cache.set('b', 2);
		cache.set('c', 3);

		// 'a' should be evicted as it's the least recently used
		assert.strictEqual(cache.get('a'), undefined);
		assert.strictEqual(cache.get('b'), 2);
		assert.strictEqual(cache.get('c'), 3);
	});

	test('Should delete by predicate', () => {
		const cache = new LRUCache<string, number>(5);
		cache.set('a1', 1);
		cache.set('a2', 2);
		cache.set('b1', 3);
		cache.set('b2', 4);

		// delete all keys starting with 'a'
		cache.deleteByPredicate(key => key.startsWith('a'));

		assert.strictEqual(cache.get('a1'), undefined);
		assert.strictEqual(cache.get('a2'), undefined);
		assert.strictEqual(cache.get('b1'), 3);
		assert.strictEqual(cache.get('b2'), 4);
	});
});

// test suite for the WorkbookCache class
suite('WorkbookCache Tests', () => {
	test('Should generate correct cache keys', () => {
		const cache = new WorkbookCache(5, 20);
		const uri = vscode.Uri.file('/path/to/file.xlsx');
		const mtime = 12345;

		const key = cache.generateKey(uri, mtime);
		assert.strictEqual(key, `${uri.toString()}-${mtime}`);

		const sheetKey = cache.generateSheetKey(key, 'Sheet1', 0);
		assert.strictEqual(sheetKey, `${key}-Sheet1-page-0`);
	});

	test('Should store and retrieve workbooks', () => {
		const cache = new WorkbookCache(5, 20);
		const mockWorkbook = { SheetNames: ['Sheet1'], Sheets: { Sheet1: {} } } as XLSX.WorkBook;

		const key = 'test-key';
		cache.setWorkbook(key, mockWorkbook);

		assert.strictEqual(cache.hasWorkbook(key), true);
		assert.deepStrictEqual(cache.getWorkbook(key), mockWorkbook);
	});

	test('Should clear caches for a URI', () => {
		const cache = new WorkbookCache(5, 20);
		const baseUri = 'file:///path/to/file.xlsx';

		// create workbook and sheet caches with the base URI
		cache.setWorkbook(`${baseUri}-123`, { SheetNames: [], Sheets: {} } as XLSX.WorkBook);
		cache.setSheet(`${baseUri}-123-Sheet1-page-0`, '<table></table>');
		cache.setWorkbook(`${baseUri}-456`, { SheetNames: [], Sheets: {} } as XLSX.WorkBook);

		// Create another cache entry with a different URI
		cache.setWorkbook('file:///other/file.xlsx-789', { SheetNames: [], Sheets: {} } as XLSX.WorkBook);

		// Clear caches for the base URI
		cache.clearCachesForUri(baseUri);

		// Check that the base URI caches are cleared but the other remains
		assert.strictEqual(cache.hasWorkbook(`${baseUri}-123`), false);
		assert.strictEqual(cache.hasSheet(`${baseUri}-123-Sheet1-page-0`), false);
		assert.strictEqual(cache.hasWorkbook(`${baseUri}-456`), false);
		assert.strictEqual(cache.hasWorkbook('file:///other/file.xlsx-789'), true);
	});
});
