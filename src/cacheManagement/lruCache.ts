/**
 * generic Least Recently Used (LRU) cache implementation
 */
export class LRUCache<K, V> {
  private cache = new Map<K, V>();
  private accessTimes = new Map<K, number>();
  private readonly maxSize: number;

  /**
   * create a new LRU cache
   * @param maxSize Maximum number of items to store in the cache
   */
  constructor(maxSize: number) {
    this.maxSize = maxSize;
  }

  /**
   * get an item from the cache
   * @param key cache key
   * @returns cached value or undefined if not found
   */
  get(key: K): V | undefined {
    if (!this.cache.has(key)) {
      return undefined;
    }
    
    // update access time
    this.accessTimes.set(key, Date.now());
    return this.cache.get(key);
  }

  /**
   * store an item in the cache
   * @param key cache key
   * @param value value to cache
   */
  set(key: K, value: V): void {
    // check if we need to make room
    if (!this.cache.has(key) && this.cache.size >= this.maxSize) {
      this.evictLeastRecentlyUsed();
    }
    
    // store the value and update access time
    this.cache.set(key, value);
    this.accessTimes.set(key, Date.now());
  }

  /**
   * check if an item exists in the cache
   * @param key The cache key
   * @returns True if the item exists
   */
  has(key: K): boolean {
    return this.cache.has(key);
  }

  /**
   * remove an item from the cache
   * @param key key to remove
   * @returns True if an item was removed
   */
  delete(key: K): boolean {
    this.accessTimes.delete(key);
    return this.cache.delete(key);
  }

  /**
   * get all keys in the cache
   * @returns Array of cache keys
   */
  keys(): K[] {
    return [...this.cache.keys()];
  }

  /**
   * get the current size of the cache
   */
  get size(): number {
    return this.cache.size;
  }

  /**
   * clear all items from the cache
   */
  clear(): void {
    this.cache.clear();
    this.accessTimes.clear();
  }

  /**
   * remove items matching a prefix
   * @param predicate Function that returns true for keys to remove
   */
  deleteByPredicate(predicate: (key: K) => boolean): void {
    for (const key of this.keys()) {
      if (predicate(key)) {
        this.delete(key);
      }
    }
  }

  /**
   * remove the least recently used item from the cache
   * @private
   */
  private evictLeastRecentlyUsed(): void {
    if (this.cache.size === 0) { return; }
    
    // find the oldest entry
    let oldestKey: K | undefined;
    let oldestTime = Infinity;
    
    for (const [key, time] of this.accessTimes.entries()) {
      if (time < oldestTime) {
        oldestTime = time;
        oldestKey = key;
      }
    }
    
    // remove the oldest entry
    if (oldestKey !== undefined) {
      console.log(`Evicting least recently used cache entry: ${String(oldestKey)}`);
      this.delete(oldestKey);
    }
  }
}