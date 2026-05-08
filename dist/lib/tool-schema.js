import { zodToJsonSchema } from "zod-to-json-schema";
function unwrapOptional(schema) {
  const def = schema._def;
  const typeName = def?.typeName;
  if (typeName === "ZodOptional" || typeName === "ZodDefault" || typeName === "ZodNullable") {
    return { inner: def.innerType, optional: true };
  }
  return { inner: schema, optional: false };
}
function describeToolSchema(tool, llmTip) {
  const params = (tool.parameters ?? []).map((p) => {
    const { inner, optional } = unwrapOptional(p.schema);
    const isPath = p.type === "Path";
    const jsonSchema = zodToJsonSchema(inner, { target: "jsonSchema7", $refStrategy: "none" });
    const { $schema: _s, ...schema } = jsonSchema;
    return {
      name: p.name,
      in: p.type,
      required: isPath || !optional,
      description: p.description,
      schema
    };
  });
  return {
    name: tool.alias,
    method: tool.method.toUpperCase(),
    path: tool.path,
    description: tool.description ?? "",
    ...llmTip ? { llmTip } : {},
    parameters: params
  };
}
function describeUtilityToolSchema(utility, ctx) {
  const schemaMap = utility.buildSchema(ctx);
  const params = Object.entries(schemaMap).map(([name, zodSchema]) => {
    const { inner, optional } = unwrapOptional(zodSchema);
    const jsonSchema = zodToJsonSchema(inner, { target: "jsonSchema7", $refStrategy: "none" });
    const { $schema: _s, ...schema } = jsonSchema;
    return {
      name,
      in: "Query",
      required: !optional,
      description: zodSchema.description,
      schema
    };
  });
  return {
    name: utility.name,
    method: utility.method,
    path: utility.path,
    description: utility.description,
    parameters: params
  };
}
export {
  describeToolSchema,
  describeUtilityToolSchema
};
