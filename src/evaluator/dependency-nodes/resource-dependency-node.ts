import { BaseEvalNode } from "./base-eval-node";

export class ResourceDependencyNode extends BaseEvalNode<{ type: "resource" }> {
  constructor(key: string) {
    super(key);
    this.setEvaluationResult({
      type: "resource",
    });
    this.resolve();
  }

  public override invalidate() {
    // Resource nodes are passive invalidation roots and always remain resolved.
  }

  public override toString(): string {
    return this.key;
  }
}
