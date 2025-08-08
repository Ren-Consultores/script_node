import { PrismaClient } from '@prisma/client';
import * as xlsx from 'xlsx';
import * as fs from 'fs';
import moment from 'moment-timezone';

const prisma = new PrismaClient();
const clientId = 30;

function getBogotaNow(): Date {
  return moment().tz('America/Bogota').toDate();
}

function convertExcelDate(excelDate: any): Date | null {
  if (!excelDate || isNaN(excelDate)) return null;
  const baseDate = new Date(Date.UTC(1899, 11, 30));
  return new Date(baseDate.getTime() + excelDate * 86400000);
}

function normalizeRowKeys<T>(row: Record<string, any>): T {
  const normalized: Record<string, any> = {};
  for (const key in row) {
    normalized[key.trim().toUpperCase()] = row[key];
  }
  return normalized as T;
}

interface ExcelRow {
  NUMERO_DE_IDENTIFICACION?: string | number;
  FECHA_DE_NACIMIENTO?: string | number;
  CELULAR?: string | number;
  ID_CREDITO?: string | number;
  VALOR_DEL_DESEMBOLSO?: string | number;
  NO_DE_CUOTAS_QUINCENALES?: string | number;
  PRIMA_TOTAL?: string | number;
  PRIMA_QUINCENAL?: string | number;
  FECHA_DESEMBOLSO?: string | number;
  VIGENCIA_POLIZA?: string | number;
  TIPO_PLAN?: string | number;
}

async function main() {
  const workbook = xlsx.readFile('./data/faltantes.xlsx');
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows = xlsx.utils.sheet_to_json<Record<string, any>>(sheet).map(row =>
    normalizeRowKeys<ExcelRow>(row)
  );

  const fails: any[] = [];
  const listComplete: any[] = [];
  let completed = 0;
  const chunkSize = 50;

  for (let i = 0; i < rows.length; i += chunkSize) {
    const chunk = rows.slice(i, i + chunkSize);

    for (const [index, row] of chunk.entries()) {
      const counter = i + index + 1;
      const now = getBogotaNow();

      try {
        const docNumber = String(row.NUMERO_DE_IDENTIFICACION).replace(/,/g, '').trim();
        if (!docNumber) {
          fails.push({ row: counter, doc_number: null, reason: 'No registra información del usuario' });
          continue;
        }

        let affiliate = await prisma.affiliates.findFirst({
          where: { doc_number: docNumber, client_id: clientId },
        });

        if (!affiliate) {
          const created = await prisma.affiliates.create({
            data: {
              doc_type: 'CC',
              doc_number: docNumber,
              client_id: clientId,
              first_name: 'VER Y MODIFICAR ESTE USUARIO',
              last_name: 'VER Y MODIFICAR ESTE USUARIO',
              birthday: convertExcelDate(row.FECHA_DE_NACIMIENTO),
              gender: 'M',
              address: 'VER Y MODIFICAR ESTE USUARIO',
              phone: '3000000',
              cellphone: String(row.CELULAR),
              email: 'VER Y MODIFICAR ESTE USUARIO',
              affiliation_type: 1,
              eps: 9,
              regional: 1,
              department: 11,
              municipality: 1,
              civil_status: 3,
              school_level: 4,
              arl: 0,
              afp: 0,
              ibc: 0,
              created_at: now,
              updated_at: now,
            },
          });

          affiliate = await prisma.affiliates.update({
            where: { id: created.id },
            data: { first_name: String(created.id), updated_at: now },
          });
        }

        const activity = await prisma.activities.create({
          data: {
            affiliate_id: affiliate.id,
            service_id: 82,
            client_id: clientId,
            state_id: 1,
            user_id: 1,
            created_at: now,
            updated_at: now,
          },
        });

        await prisma.life_debtor_insurances.create({
          data: {
            activity_id: activity.id,
            id_number: Number(docNumber),
            id_credit: String(row.ID_CREDITO),
            cellphone: String(row.CELULAR),
            disbursement_value: String(row.VALOR_DEL_DESEMBOLSO),
            quantity_quotes: String(row.NO_DE_CUOTAS_QUINCENALES),
            disbursement_date: convertExcelDate(row.FECHA_DESEMBOLSO),
            secure_prime_value: Number(row.PRIMA_TOTAL),
            biweekly_premium: String(row.PRIMA_QUINCENAL),
            policy_validity: convertExcelDate(row.VIGENCIA_POLIZA),
            plan_type: row.TIPO_PLAN ? String(row.TIPO_PLAN) : undefined,
            created_at: now,
            updated_at: now,
          },
        });

        await prisma.activity_actions.create({
          data: {
            activity_id: activity.id,
            action_id: 256,
            old_user_id: 1,
            new_user_id: 1,
            description: 'Se reporta colocación desde cargue masivo',
            old_state_id: 1,
            new_state_id: 113,
            author_id: 1,
            created_at: now,
            updated_at: now,
          },
        });

        const activityAction2 = await prisma.activity_actions.create({
          data: {
            activity_id: activity.id,
            action_id: 258,
            old_user_id: 1,
            new_user_id: 1,
            description: 'Se genera certificado kit de bienvenida pólizas vida grupo',
            old_state_id: 113,
            new_state_id: 114,
            author_id: 1,
            created_at: now,
            updated_at: now,
          },
        });

        await prisma.life_debtor_insurance_without_reports.create({
          data: {
            activity_id: activity.id,
            id_credit: Number(row.ID_CREDITO),
            doc_number: docNumber,
            cellphone: String(row.CELULAR),
            client: String(clientId),
            activity_2_id: String(activityAction2.id),
            affiliate_id: String(affiliate.id),
            req: JSON.stringify(row),
            created_at: now,
            updated_at: now,
          },
        });

        await prisma.activities.update({
          where: { id: activity.id },
          data: {
            user_id: 1,
            state_id: 114,
            updated_at: now,
          },
        });

        listComplete.push({
          service_id: activity.id,
          nro_documento: docNumber,
          nombre_afiliado: `${affiliate.first_name} ${affiliate.last_name}`,
          nro_credit: activity.id,
        });

        completed++;
      } catch (error) {
        const errorMessage = error instanceof Error ? error.message : String(error);
        console.error(`❌ Error en fila ${counter} - Cédula: ${row.NUMERO_DE_IDENTIFICACION}:`, errorMessage);
        fails.push({
          row: counter,
          doc_number: row.NUMERO_DE_IDENTIFICACION ?? 'N/D',
          error: errorMessage,
        });
      }
    }
  }

  console.log('\n✅ Proceso de cargue finalizado.');
  console.log('✔ Completados:', completed);
  console.log('❌ Fallidos:', fails.length);
  console.table(listComplete);

  if (fails.length > 0) {
    fs.writeFileSync('./fallidos.json', JSON.stringify(fails, null, 2));
    console.log('📁 Archivo fallidos.json generado con los errores.');
  }
}

main()
  .catch(e => console.error('🔥 Error crítico:', e))
  .finally(() => prisma.$disconnect());
