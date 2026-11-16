export class DeviceOperationQueue {
  private tail = Promise.resolve();
  private sequence = 0;

  round(operation: (sequence: number) => Promise<void>): Promise<void> {
    return this.enqueue(() => operation(++this.sequence));
  }

  properties(operation: () => Promise<void>): Promise<void> {
    return this.enqueue(operation);
  }

  private enqueue(operation: () => Promise<void>): Promise<void> {
    const result = this.tail.then(operation);
    this.tail = result;
    return result;
  }
}
