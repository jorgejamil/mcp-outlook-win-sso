#!/usr/bin/env node
import { Server } from '@modelcontextprotocol/sdk/server/index.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import {
  CallToolRequestSchema,
  ListToolsRequestSchema,
} from '@modelcontextprotocol/sdk/types.js';
import { Client } from '@microsoft/microsoft-graph-client';
import { TokenCredentialAuthenticationProvider } from '@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials/index.js';
import { InteractiveBrowserCredential } from '@azure/identity';
import open from 'open';
import express from 'express';
import { promises as fs } from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

class OutlookMCPServer {
  constructor() {
    this.server = new Server(
      {
        name: 'outlook-mcp-server',
        version: '1.0.0',
      },
      {
        capabilities: {
          tools: {},
        },
      }
    );

    this.client = null;
    this.setupHandlers();
  }

  async initializeClient() {
    if (this.client) return;

    try {
      // Carregar configuração
      const configPath = path.join(__dirname, 'config.json');
      const config = JSON.parse(await fs.readFile(configPath, 'utf8'));

      // Configurar credencial interativa do browser
      const credential = new InteractiveBrowserCredential({
        clientId: config.clientId,
        tenantId: config.tenantId,
        redirectUri: 'http://localhost:3000/auth/callback',
      });

      // Criar provider de autenticação
      const authProvider = new TokenCredentialAuthenticationProvider(credential, {
        scopes: ['https://graph.microsoft.com/.default'],
      });

      // Inicializar cliente do Microsoft Graph
      this.client = Client.initWithMiddleware({
        authProvider: authProvider,
      });

      console.error('Cliente Microsoft Graph inicializado com sucesso');
    } catch (error) {
      console.error('Erro ao inicializar cliente:', error);
      throw error;
    }
  }

  setupHandlers() {
    this.server.setRequestHandler(ListToolsRequestSchema, async () => ({
      tools: [
        {
          name: 'list_emails',
          description: 'Lista emails da caixa de entrada',
          inputSchema: {
            type: 'object',
            properties: {
              folder: {
                type: 'string',
                description: 'Pasta do email (inbox, sent, drafts)',
                default: 'inbox',
              },
              limit: {
                type: 'number',
                description: 'Número máximo de emails',
                default: 10,
              },
              search: {
                type: 'string',
                description: 'Termo de busca opcional',
              },
            },
          },
        },
        {
          name: 'read_email',
          description: 'Lê o conteúdo completo de um email',
          inputSchema: {
            type: 'object',
            properties: {
              emailId: {
                type: 'string',
                description: 'ID do email',
              },
            },
            required: ['emailId'],
          },
        },
        {
          name: 'send_email',
          description: 'Envia um novo email',
          inputSchema: {
            type: 'object',
            properties: {
              to: {
                type: 'array',
                items: { type: 'string' },
                description: 'Lista de destinatários',
              },
              subject: {
                type: 'string',
                description: 'Assunto do email',
              },
              body: {
                type: 'string',
                description: 'Corpo do email',
              },
              cc: {
                type: 'array',
                items: { type: 'string' },
                description: 'Lista de emails em cópia',
              },
              isHtml: {
                type: 'boolean',
                description: 'Se o corpo é HTML',
                default: false,
              },
            },
            required: ['to', 'subject', 'body'],
          },
        },
        {
          name: 'list_calendar_events',
          description: 'Lista eventos do calendário',
          inputSchema: {
            type: 'object',
            properties: {
              startDate: {
                type: 'string',
                description: 'Data inicial (ISO 8601)',
              },
              endDate: {
                type: 'string',
                description: 'Data final (ISO 8601)',
              },
              limit: {
                type: 'number',
                description: 'Número máximo de eventos',
                default: 20,
              },
            },
          },
        },
        {
          name: 'create_calendar_event',
          description: 'Cria um novo evento no calendário',
          inputSchema: {
            type: 'object',
            properties: {
              subject: {
                type: 'string',
                description: 'Título do evento',
              },
              start: {
                type: 'string',
                description: 'Data/hora de início (ISO 8601)',
              },
              end: {
                type: 'string',
                description: 'Data/hora de fim (ISO 8601)',
              },
              body: {
                type: 'string',
                description: 'Descrição do evento',
              },
              location: {
                type: 'string',
                description: 'Local do evento',
              },
              attendees: {
                type: 'array',
                items: { type: 'string' },
                description: 'Lista de emails dos participantes',
              },
              isOnline: {
                type: 'boolean',
                description: 'Se é um evento online',
                default: false,
              },
            },
            required: ['subject', 'start', 'end'],
          },
        },
      ],
    }));

    this.server.setRequestHandler(CallToolRequestSchema, async (request) => {
      await this.initializeClient();

      const { name, arguments: args } = request.params;

      try {
        switch (name) {
          case 'list_emails':
            return await this.listEmails(args);
          case 'read_email':
            return await this.readEmail(args);
          case 'send_email':
            return await this.sendEmail(args);
          case 'list_calendar_events':
            return await this.listCalendarEvents(args);
          case 'create_calendar_event':
            return await this.createCalendarEvent(args);
          default:
            throw new Error(`Ferramenta desconhecida: ${name}`);
        }
      } catch (error) {
        return {
          content: [
            {
              type: 'text',
              text: `Erro: ${error.message}`,
            },
          ],
        };
      }
    });
  }

  async listEmails({ folder = 'inbox', limit = 10, search }) {
    try {
      let endpoint = '/me/mailFolders/' + folder + '/messages';
      let queryParams = [`$top=${limit}`, '$orderby=receivedDateTime desc'];

      if (search) {
        queryParams.push(`$search="${search}"`);
      }

      const response = await this.client
        .api(endpoint)
        .query(queryParams.join('&'))
        .get();

      const emails = response.value.map((email) => ({
        id: email.id,
        subject: email.subject,
        from: email.from?.emailAddress?.address,
        received: email.receivedDateTime,
        hasAttachments: email.hasAttachments,
        preview: email.bodyPreview,
      }));

      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify(emails, null, 2),
          },
        ],
      };
    } catch (error) {
      throw new Error(`Erro ao listar emails: ${error.message}`);
    }
  }

  async readEmail({ emailId }) {
    try {
      const email = await this.client
        .api(`/me/messages/${emailId}`)
        .select('subject,body,from,to,cc,receivedDateTime,hasAttachments')
        .get();

      let attachments = [];
      if (email.hasAttachments) {
        const attachmentsResponse = await this.client
          .api(`/me/messages/${emailId}/attachments`)
          .get();
        attachments = attachmentsResponse.value.map((att) => ({
          name: att.name,
          size: att.size,
          contentType: att.contentType,
        }));
      }

      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify(
              {
                ...email,
                attachments,
              },
              null,
              2
            ),
          },
        ],
      };
    } catch (error) {
      throw new Error(`Erro ao ler email: ${error.message}`);
    }
  }

  async sendEmail({ to, subject, body, cc = [], isHtml = false }) {
    try {
      const message = {
        subject,
        body: {
          contentType: isHtml ? 'HTML' : 'Text',
          content: body,
        },
        toRecipients: to.map((email) => ({
          emailAddress: { address: email },
        })),
      };

      if (cc.length > 0) {
        message.ccRecipients = cc.map((email) => ({
          emailAddress: { address: email },
        }));
      }

      await this.client.api('/me/sendMail').post({
        message,
        saveToSentItems: true,
      });

      return {
        content: [
          {
            type: 'text',
            text: 'Email enviado com sucesso!',
          },
        ],
      };
    } catch (error) {
      throw new Error(`Erro ao enviar email: ${error.message}`);
    }
  }

  async listCalendarEvents({ startDate, endDate, limit = 20 }) {
    try {
      let queryParams = [`$top=${limit}`];

      if (startDate && endDate) {
        queryParams.push(
          `$filter=start/dateTime ge '${startDate}' and end/dateTime le '${endDate}'`
        );
      }

      const response = await this.client
        .api('/me/events')
        .query(queryParams.join('&'))
        .orderby('start/dateTime')
        .get();

      const events = response.value.map((event) => ({
        id: event.id,
        subject: event.subject,
        start: event.start.dateTime,
        end: event.end.dateTime,
        location: event.location?.displayName,
        isOnlineMeeting: event.isOnlineMeeting,
        organizer: event.organizer?.emailAddress?.address,
      }));

      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify(events, null, 2),
          },
        ],
      };
    } catch (error) {
      throw new Error(`Erro ao listar eventos: ${error.message}`);
    }
  }

  async createCalendarEvent({
    subject,
    start,
    end,
    body = '',
    location = '',
    attendees = [],
    isOnline = false,
  }) {
    try {
      const event = {
        subject,
        body: {
          contentType: 'HTML',
          content: body,
        },
        start: {
          dateTime: start,
          timeZone: 'America/Sao_Paulo',
        },
        end: {
          dateTime: end,
          timeZone: 'America/Sao_Paulo',
        },
        location: {
          displayName: location,
        },
        isOnlineMeeting: isOnline,
      };

      if (attendees.length > 0) {
        event.attendees = attendees.map((email) => ({
          emailAddress: { address: email },
          type: 'required',
        }));
      }

      const response = await this.client.api('/me/events').post(event);

      return {
        content: [
          {
            type: 'text',
            text: `Evento criado com sucesso! ID: ${response.id}`,
          },
        ],
      };
    } catch (error) {
      throw new Error(`Erro ao criar evento: ${error.message}`);
    }
  }

  async run() {
    const transport = new StdioServerTransport();
    await this.server.connect(transport);
    console.error('Servidor MCP Outlook iniciado');
  }
}

const server = new OutlookMCPServer();
server.run().catch(console.error);