initDb()
  .then(() => {
    app.listen(PORT, () => {
      console.log(`Financial importer running on http://localhost:${PORT}/`);
    });
  })
  .catch((err) => {
    console.error('Database initialization failed:', err);
    process.exit(1);
  });
