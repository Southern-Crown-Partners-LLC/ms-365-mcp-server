import { z } from "zod";
function makeApi(endpoints) {
  return endpoints;
}
class Zodios {
  constructor(baseUrlOrEndpoints, endpoints, options) {
    if (typeof baseUrlOrEndpoints === "string") {
      throw new Error("No such hack");
    }
    this.endpoints = baseUrlOrEndpoints.map((endpoint) => {
      endpoint.parameters = endpoint.parameters || [];
      for (const parameter of endpoint.parameters) {
        parameter.name = parameter.name.replace(/[$_]+/g, "");
      }
      const pathParamRegex = /:([a-zA-Z0-9]+)/g;
      const pathParams = [];
      let match;
      while ((match = pathParamRegex.exec(endpoint.path)) !== null) {
        pathParams.push(match[1]);
      }
      for (const pathParam of pathParams) {
        const paramExists = endpoint.parameters.some(
          (param) => param.name === pathParam || param.name === pathParam.replace(/[$_]+/g, "")
        );
        if (!paramExists) {
          const newParam = {
            name: pathParam,
            type: "Path",
            schema: z.string().describe(`Path parameter: ${pathParam}`),
            description: `Path parameter: ${pathParam}`
          };
          endpoint.parameters.push(newParam);
        }
      }
      return endpoint;
    });
  }
}
export {
  Zodios,
  makeApi
};
