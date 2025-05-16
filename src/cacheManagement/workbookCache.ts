import * as vscode from 'vscode';
import * as XLSX from 'xlsx';
import { LRUCache } from './lruCache';

/**
 * specialized cache for SheetJS workbooks with URI-based keys
 */
export class WorkbookCache {
  private workbookCache: LRUCache<string, XLSX.WorkBook>;
  private sheetCache: LRUCache<string, string>;
  
  /**
   * create a new workbook cache
   * @param maxWorkbooks Maximum number of workbooks to cache
   * @param maxSheets Maximum number of sheet HTML to cache
   */
  constructor(maxWorkbooks: number = 10, maxSheets: number = 255) {
    this.workbookCache = new LRUCache<string, XLSX.WorkBook>(maxWorkbooks);
    this.sheetCache = new LRUCache<string, string>(maxSheets);
  }
  
  /**
   * generate a cache key for a document
   * @param uri Document URI
   * @param mtime Modification time
   */
  generateKey(uri: vscode.Uri, mtime: number): string {
    return `${uri.toString()}-${mtime}`;
  }
  
  /**
   * get a workbook from the cache
   * @param key Cache key
   */
  getWorkbook(key: string): XLSX.WorkBook | undefined {
    return this.workbookCache.get(key);
  }
  
  /**
   * store a workbook in the cache
   * @param key Cache key
   * @param workbook Workbook to cache
   */
  setWorkbook(key: string, workbook: XLSX.WorkBook): void {
    this.workbookCache.set(key, workbook);
  }
  
  /**
   * check if a workbook exists in the cache
   * @param key Cache key
   */
  hasWorkbook(key: string): boolean {
    return this.workbookCache.has(key);
  }
  
  /**
   * get sheet HTML from the cache
   * @param key Sheet cache key
   */
  getSheet(key: string): string | undefined {
    return this.sheetCache.get(key);
  }
  
  /**
   * store sheet HTML in the cache
   * @param key Sheet cache key
   * @param html Sheet HTML
   */
  setSheet(key: string, html: string): void {
    this.sheetCache.set(key, html);
  }
  
  /**
   * Check if sheet HTML exists in the cache
   * @param key Sheet cache key
   */
  hasSheet(key: string): boolean {
    return this.sheetCache.has(key);
  }
  
  /**
   * generate a sheet cache key
   * @param baseKey Base workbook key
   * @param sheetName Sheet name
   * @param page Page number
   */
  generateSheetKey(baseKey: string, sheetName: string, page: number): string {
    return `${baseKey}-${sheetName}-page-${page}`;
  }
  
  /**
   * clear all caches for a specific URI
   * @param uriString URI string prefix to clear
   */
  clearCachesForUri(uriString: string): void {
    console.log(`Clearing caches for ${uriString}`);
    
    // clear workbook cache entries for this URI
    this.workbookCache.deleteByPredicate(key => key.startsWith(uriString));
    
    // clear sheet cache entries for this URI
    this.sheetCache.deleteByPredicate(key => key.startsWith(uriString));
  }
}