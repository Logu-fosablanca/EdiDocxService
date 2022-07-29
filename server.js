import express  from "express";
import { createServer } from "http";
import bodyparser from "body-parser";

let app = express();
let port = process.env.PORT || 3000;
let server = createServer(app);
server.listen(port);
console.log("Server listening on port " + port);


import {getDocx} from "./docx.js";

app.use(bodyparser.json());




app.post('/api/create/docx',getDocx);