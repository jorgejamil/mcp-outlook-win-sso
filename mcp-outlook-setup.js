import { promises as fs } from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import readline from 'readline';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const rl = readline.createInterface({
  input: process.stdin,
  output: process.stdout
});

const question = (query) => new Promise((resolve) => rl.question(query, resolve));

async function setup() {
  console.log('=== Configuração do Servidor MCP Outlook ===\n');
  
  console.log('Para configurar este servidor, você precisará:');
  console.log('1. Um Azure App Registration');
  console.log('2. Client ID e Tenant ID da aplicação');
  console.log('3. Permissões configuradas no Azure AD\n');
  
  console.log('Siga estas etapas no Azure Portal:');
  console.log('1. Acesse https://portal.azure.com');
  console.log('2. Vá para Azure Active Directory > App registrations');
  console.log('3. Clique em "New registration"');
  console.log('4. Nome: "MCP Outlook Server"');
  console.log('5. Supported account types: "Accounts in this organizational directory only"');
  console.log('6. Redirect URI: Web - http://localhost:3000/auth/callback');
  console.log('7. Após criar, anote o Application (client) ID e Directory (tenant) ID\n');
  
  console.log('8. Em "API permissions", adicione:');
  console.log('   - Microsoft Graph > Delegated permissions:');
  console.log('     * Mail.Read');
  console.log('     * Mail.ReadWrite');
  console.log('     * Mail.Send');
  console.log('     * Calendars.Read');
  console.log('     * Calendars.ReadWrite');
  console.log('     * User.Read');
  console.log('9. Clique em "Grant admin consent" (pode precisar de um admin)\n');
  
  console.log('10. Em "Authentication":');
  console.log('    - Certifique-se que "http://localhost:3000/auth/callback" está nas Redirect URIs');
  console.log('    - Em "Implicit grant and hybrid flows", marque:');
  console.log('      * Access tokens');
  console.log('      * ID tokens\n');
  
  const proceed = await question('Você já completou essas etapas? (s/n): ');
  
  if (proceed.toLowerCase() !== 's') {
    console.log('\nComplete as etapas acima e execute novamente este script.');
    process.exit(0);
  }
  
  console.log('\n--- Configuração ---\n');
  
  const clientId = await question('Digite o Client ID: ');
  const tenantId = await question('Digite o Tenant ID: ');
  
  const config = {
    clientId: clientId.trim(),
    tenantId: tenantId.trim(),
    redirectUri: 'http://localhost:3000/auth/callback',
    scopes: [
      'https://graph.microsoft.com/Mail.Read',
      'https://graph.microsoft.com/Mail.ReadWrite',
      'https://graph.microsoft.com/Mail.Send',
      'https://graph.microsoft.com/Calendars.Read',
      'https://graph.microsoft.com/Calendars.ReadWrite',
      'https://graph.microsoft.com/User.Read'
    ]
  };
  
  const configPath = path.join(__dirname, 'config.json');
  await fs.writeFile(configPath, JSON.stringify(config, null, 2));
  
  console.log('\n✅ Configuração salva com sucesso!');
  console.log(`📁 Arquivo: ${configPath}`);
  
  console.log('\n--- Próximos passos ---\n');
  console.log('1. Instale as dependências:');
  console.log('   npm install\n');
  
  console.log('2. Configure o Claude Desktop:');
  console.log('   Adicione ao arquivo de configuração do Claude Desktop:');
  console.log('   (normalmente em %APPDATA%/Claude/claude_desktop_config.json)\n');
  
  const claudeConfig = {
    "mcpServers": {
      "outlook": {
        "command": "npx",
        "args": ["-y", "outlook-mcp-server"],
        "env": {}
      }
    }
  };
  
  console.log(JSON.stringify(claudeConfig, null, 2));
  
  console.log('\n3. Reinicie o Claude Desktop');
  console.log('\n4. Na primeira execução, uma janela do navegador abrirá para autenticação');
  console.log('   Faça login com sua conta corporativa do Office 365\n');
  
  rl.close();
}

setup().catch(console.error);