


import { Configuration, OpenAIApi } from "openai";

const TEMPLATE = './PRESENTATION_TITLE.pptx';
const OUTPUT = './output.pptx';

const configuration = new Configuration({
  apiKey: "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",
});
const openai = new OpenAIApi(configuration);
const animal = `Create a ppt with the following description, output in json format: I am a programmer, generate a year-end summary ppt`;
const completion = await openai.createCompletion({
  model: "text-davinci-003",
  prompt: animal,
  max_tokens: Math.min(1000),
  temperature: 0,
});

console.log('completion: ', completion);

