import express from "express";
import path from "path";
import cors from "cors";
import { createServer as createViteServer } from "vite";
import { db } from "./src/db";

async function startServer() {
  const app = express();
  const PORT = 3000;

  app.use(cors());
  app.use(express.json({ limit: "50mb" }));

  // API endpoints for records
  app.get("/api/records", (req, res) => {
    try {
      const stmt = db.prepare("SELECT * FROM records");
      const records = stmt.all();
      res.json(records);
    } catch (error) {
      console.error(error);
      res.status(500).json({ error: "Failed to fetch records" });
    }
  });

  app.post("/api/records/bulk", (req, res) => {
    try {
      const records = req.body.records;
      if (!Array.isArray(records)) {
        return res.status(400).json({ error: "Invalid data format" });
      }

      const stmt = db.prepare(`
        INSERT OR REPLACE INTO records (
          id, admissao, horaClassificacao, atendimento, fimAtendimento, classificacao,
          rawAdmissao, rawHoraClassificacao, rawAtendimento, rawFimAtendimento
        ) VALUES (
          @id, @admissao, @horaClassificacao, @atendimento, @fimAtendimento, @classificacao,
          @rawAdmissao, @rawHoraClassificacao, @rawAtendimento, @rawFimAtendimento
        )
      `);

      const insertMany = db.transaction((rows) => {
        for (const row of rows) {
          stmt.run(row);
        }
      });

      insertMany(records);
      res.json({ message: "Records inserted successfully", count: records.length });
    } catch (error) {
      console.error(error);
      res.status(500).json({ error: "Failed to insert records" });
    }
  });

  app.put("/api/records/:id", (req, res) => {
    try {
      const id = req.params.id;
      const { 
        admissao, horaClassificacao, atendimento, fimAtendimento, classificacao,
        rawAdmissao, rawHoraClassificacao, rawAtendimento, rawFimAtendimento 
      } = req.body;

      const stmt = db.prepare(`
        UPDATE records SET 
          admissao = @admissao,
          horaClassificacao = @horaClassificacao,
          atendimento = @atendimento,
          fimAtendimento = @fimAtendimento,
          classificacao = @classificacao,
          rawAdmissao = @rawAdmissao,
          rawHoraClassificacao = @rawHoraClassificacao,
          rawAtendimento = @rawAtendimento,
          rawFimAtendimento = @rawFimAtendimento
        WHERE id = @id
      `);

      stmt.run({
        id, admissao, horaClassificacao, atendimento, fimAtendimento, classificacao,
        rawAdmissao, rawHoraClassificacao, rawAtendimento, rawFimAtendimento
      });
      
      res.json({ message: "Record updated successfully" });
    } catch (error) {
      console.error(error);
      res.status(500).json({ error: "Failed to update record" });
    }
  });

  app.delete("/api/records/bulk-delete", (req, res) => {
    try {
      const ids = req.body.ids;
      if (!Array.isArray(ids)) {
        return res.status(400).json({ error: "Invalid data format" });
      }

      const stmt = db.prepare("DELETE FROM records WHERE id = ?");
      
      const deleteMany = db.transaction((idList) => {
        for (const id of idList) {
          stmt.run(id);
        }
      });

      deleteMany(ids);
      res.json({ message: "Records deleted successfully" });
    } catch (error) {
      console.error(error);
      res.status(500).json({ error: "Failed to delete records" });
    }
  });

  app.delete("/api/records/:id", (req, res) => {
    try {
      const stmt = db.prepare("DELETE FROM records WHERE id = ?");
      stmt.run(req.params.id);
      res.json({ message: "Record deleted successfully" });
    } catch (error) {
      console.error(error);
      res.status(500).json({ error: "Failed to delete record" });
    }
  });

  // Exemplo de rota da API para testar o banco de dados
  app.get("/api/health", (req, res) => {
    res.json({ 
      status: "ok", 
      db: db.open ? "connected" : "disconnected",
      message: "Banco de dados better-sqlite3 inicializado com sucesso"
    });
  });

  // Vite middleware for development
  if (process.env.NODE_ENV !== "production") {
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: "spa",
    });
    app.use(vite.middlewares);
  } else {
    // Modo de produção: serve os arquivos estáticos compilados do React
    const distPath = path.join(process.cwd(), 'dist');
    app.use(express.static(distPath));
    app.get('*', (req, res) => {
      res.sendFile(path.join(distPath, 'index.html'));
    });
  }

  app.listen(PORT, "0.0.0.0", () => {
    console.log(`Servidor rodando na porta ${PORT}`);
  });
}

startServer();
