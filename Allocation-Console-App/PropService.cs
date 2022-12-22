using Allocation_Console_App.Entities;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Allocation_Console_App
{
    public class PropService
    {

        public List<Table> CreatePropTables(bool isTest = false)
        {
            int[] TableSeatCounts = { 78, 78, 78, 72, 72 };

            int[] testTable = { 20 };


            List<Table> tables = new();

            if (isTest)
            {
                for (int i = 1; i <= testTable.Length; i++)
                {
                    Table table = new Table();
                    table.TableNumber = i * 100;
                    table.Seats = CreatePropSeats(table, testTable[i - 1]);
                    tables.Add(table);
                }
                SetHardcodedSeatValues(tables, true);
            }
            else
            {
                for (int i = 1; i <= TableSeatCounts.Length; i++)
                {
                    Table table = new Table();
                    table.TableNumber = i * 100;
                    table.Seats = CreatePropSeats(table, TableSeatCounts[i - 1]);
                    tables.Add(table);
                }
                SetHardcodedSeatValues(tables);
            }
            return tables;
        }

        public List<Seat> CreatePropSeats(Table parentTable, int seatCount)
        {
            List<Seat> seats = new();
            if (parentTable == null)
            {
                return seats;
            }

            for (int i = 0; i < seatCount; i++)
            {
                Seat seat = new Seat();
                seat.SeatNumber = parentTable.TableNumber + i;
                seat.Parent = parentTable;
                seat.Score = 1;
                seat.Occupied = false;
                seats.Add(seat);
            }

            return seats;
        }

        private void SetHardcodedSeatValues(List<Table> tables, bool isTest = false)
        {
            if (isTest)
            {
                foreach(var seat in tables[0].Seats)
                {
                    switch (seat.SeatNumber)
                    {
                        case 100:
                            seat.Score = 4;
                            break;
                        case 101:
                            seat.Score = 5;
                            break;
                        case 102:
                            seat.Score = 6;
                            break;
                        case 103:
                            seat.Score = 7;
                            break;
                        case 104:
                            seat.Score = 8;
                            break;
                        case 105:
                            seat.Score = 9;
                            break;
                        case 106:
                            seat.Score = 10;
                            break;
                        case 107:
                            seat.Score = 9;
                            break;
                        case 108:
                            seat.Score = 8;
                            break;
                        case 109:
                            seat.Score = 7;
                            break;
                        case 110:
                            seat.Score = 6;
                            break;
                        case 111:
                            seat.Score = 5;
                            break;
                        case 112:
                            seat.Score = 4;
                            break;
                        case 113:
                            seat.Score = 3;
                            break;
                        case 114:
                            seat.Score = 2;
                            break;
                        case 115:
                            seat.Score = 1;
                            break;
                        case 116:
                            seat.Score = 1;
                            break;
                        case 117:
                            seat.Score = 1;
                            break;
                        case 118:
                            seat.Score = 1;
                            break;
                        case 119:
                            seat.Score = 1;
                            break;
                        default:
                            seat.Score = 1;
                            break;
                    }
                }
            }
            else
            {

                foreach (var seat in tables[0].Seats)
                {
                    switch (seat.SeatNumber)
                    {
                        case 100:
                            seat.Score = 5;
                            break;
                        case 101:
                            seat.Score = 5;
                            break;
                        case 102:
                            seat.Score = 5;
                            break;
                        case 103:
                            seat.Score = 5;
                            break;
                        case 104:
                            seat.Score = 10;
                            break;
                        case 105:
                            seat.Score = 10;
                            break;
                        case 106:
                            seat.Score = 15;
                            break;
                        case 107:
                            seat.Score = 15;
                            break;
                        case 108:
                            seat.Score = 20;
                            break;
                        case 109:
                            seat.Score = 20;
                            break;
                        case 110:
                            seat.Score = 25;
                            break;
                        case 111:
                            seat.Score = 25;
                            break;
                        case 112:
                            seat.Score = 30;
                            break;
                        case 113:
                            seat.Score = 30;
                            break;
                        case 114:
                            seat.Score = 35;
                            break;
                        case 115:
                            seat.Score = 35;
                            break;
                        case 116:
                            seat.Score = 40;
                            break;
                        case 117:
                            seat.Score = 40;
                            break;
                        case 118:
                            seat.Score = 45;
                            break;
                        case 119:
                            seat.Score = 45;
                            break;
                        case 120:
                            seat.Score = 50;
                            break;
                        case 121:
                            seat.Score = 50;
                            break;
                        case 122:
                            seat.Score = 45;
                            break;
                        case 123:
                            seat.Score = 45;
                            break;
                        case 124:
                            seat.Score = 40;
                            break;
                        case 125:
                            seat.Score = 40;
                            break;
                        case 126:
                            seat.Score = 35;
                            break;
                        case 127:
                            seat.Score = 35;
                            break;
                        case 128:
                            seat.Score = 30;
                            break;
                        case 129:
                            seat.Score = 30;
                            break;
                        case 130:
                            seat.Score = 25;
                            break;
                        case 131:
                            seat.Score = 25;
                            break;
                        case 132:
                            seat.Score = 20;
                            break;
                        case 133:
                            seat.Score = 20;
                            break;
                        case 134:
                            seat.Score = 15;
                            break;
                        case 135:
                            seat.Score = 15;
                            break;
                        case 136:
                            seat.Score = 10;
                            break;
                        case 137:
                            seat.Score = 10;
                            break;
                        default:
                            seat.Score = 5;
                            break;
                    }
                }

                foreach (var seat in tables[1].Seats)
                {
                    switch (seat.SeatNumber)
                    {
                        case 200:
                            seat.Score = 25;
                            break;
                        case 201:
                            seat.Score = 25;
                            break;
                        case 202:
                            seat.Score = 30;
                            break;
                        case 203:
                            seat.Score = 30;
                            break;
                        case 204:
                            seat.Score = 35;
                            break;
                        case 205:
                            seat.Score = 35;
                            break;
                        case 206:
                            seat.Score = 40;
                            break;
                        case 207:
                            seat.Score = 40;
                            break;
                        case 208:
                            seat.Score = 45;
                            break;
                        case 209:
                            seat.Score = 45;
                            break;
                        case 210:
                            seat.Score = 50;
                            break;
                        case 211:
                            seat.Score = 50;
                            break;
                        case 212:
                            seat.Score = 55;
                            break;
                        case 213:
                            seat.Score = 55;
                            break;
                        case 214:
                            seat.Score = 60;
                            break;
                        case 215:
                            seat.Score = 60;
                            break;
                        case 216:
                            seat.Score = 65;
                            break;
                        case 217:
                            seat.Score = 65;
                            break;
                        case 218:
                            seat.Score = 70;
                            break;
                        case 219:
                            seat.Score = 70;
                            break;
                        case 220:
                            seat.Score = 75;
                            break;
                        case 221:
                            seat.Score = 75;
                            break;
                        case 222:
                            seat.Score = 70;
                            break;
                        case 223:
                            seat.Score = 70;
                            break;
                        case 224:
                            seat.Score = 65;
                            break;
                        case 225:
                            seat.Score = 65;
                            break;
                        case 226:
                            seat.Score = 60;
                            break;
                        case 227:
                            seat.Score = 60;
                            break;
                        case 228:
                            seat.Score = 55;
                            break;
                        case 229:
                            seat.Score = 55;
                            break;
                        case 230:
                            seat.Score = 50;
                            break;
                        case 231:
                            seat.Score = 50;
                            break;
                        case 232:
                            seat.Score = 45;
                            break;
                        case 233:
                            seat.Score = 45;
                            break;
                        case 234:
                            seat.Score = 40;
                            break;
                        case 235:
                            seat.Score = 40;
                            break;
                        case 236:
                            seat.Score = 35;
                            break;
                        case 237:
                            seat.Score = 35;
                            break;
                        case 238:
                            seat.Score = 30;
                            break;
                        case 239:
                            seat.Score = 30;
                            break;
                        case 240:
                            seat.Score = 25;
                            break;
                        case 241:
                            seat.Score = 25;
                            break;
                        case 242:
                            seat.Score = 20;
                            break;
                        case 243:
                            seat.Score = 20;
                            break;
                        case 244:
                            seat.Score = 15;
                            break;
                        case 245:
                            seat.Score = 15;
                            break;
                        case 246:
                            seat.Score = 10;
                            break;
                        case 247:
                            seat.Score = 10;
                            break;
                        default:
                            seat.Score = 5;
                            break;
                    }
                }

                foreach (var seat in tables[2].Seats)
                {
                    switch (seat.SeatNumber)
                    {
                        case 300:
                            seat.Score = 50;
                            break;
                        case 301:
                            seat.Score = 50;
                            break;
                        case 302:
                            seat.Score = 55;
                            break;
                        case 303:
                            seat.Score = 55;
                            break;
                        case 304:
                            seat.Score = 60;
                            break;
                        case 305:
                            seat.Score = 60;
                            break;
                        case 306:
                            seat.Score = 65;
                            break;
                        case 307:
                            seat.Score = 65;
                            break;
                        case 308:
                            seat.Score = 70;
                            break;
                        case 309:
                            seat.Score = 70;
                            break;
                        case 310:
                            seat.Score = 75;
                            break;
                        case 311:
                            seat.Score = 75;
                            break;
                        case 312:
                            seat.Score = 80;
                            break;
                        case 313:
                            seat.Score = 80;
                            break;
                        case 314:
                            seat.Score = 85;
                            break;
                        case 315:
                            seat.Score = 85;
                            break;
                        case 316:
                            seat.Score = 90;
                            break;
                        case 317:
                            seat.Score = 90;
                            break;
                        case 318:
                            seat.Score = 95;
                            break;
                        case 319:
                            seat.Score = 95;
                            break;
                        case 320:
                            seat.Score = 100;
                            break;
                        case 321:
                            seat.Score = 100;
                            break;
                        case 322:
                            seat.Score = 95;
                            break;
                        case 323:
                            seat.Score = 95;
                            break;
                        case 324:
                            seat.Score = 90;
                            break;
                        case 325:
                            seat.Score = 90;
                            break;
                        case 326:
                            seat.Score = 85;
                            break;
                        case 327:
                            seat.Score = 85;
                            break;
                        case 328:
                            seat.Score = 80;
                            break;
                        case 329:
                            seat.Score = 80;
                            break;
                        case 330:
                            seat.Score = 75;
                            break;
                        case 331:
                            seat.Score = 75;
                            break;
                        case 332:
                            seat.Score = 70;
                            break;
                        case 333:
                            seat.Score = 70;
                            break;
                        case 334:
                            seat.Score = 65;
                            break;
                        case 335:
                            seat.Score = 65;
                            break;
                        case 336:
                            seat.Score = 60;
                            break;
                        case 337:
                            seat.Score = 60;
                            break;
                        case 338:
                            seat.Score = 55;
                            break;
                        case 339:
                            seat.Score = 55;
                            break;
                        case 340:
                            seat.Score = 50;
                            break;
                        case 341:
                            seat.Score = 50;
                            break;
                        case 342:
                            seat.Score = 45;
                            break;
                        case 343:
                            seat.Score = 45;
                            break;
                        case 344:
                            seat.Score = 40;
                            break;
                        case 345:
                            seat.Score = 40;
                            break;
                        case 346:
                            seat.Score = 35;
                            break;
                        case 347:
                            seat.Score = 35;
                            break;
                        case 348:
                            seat.Score = 30;
                            break;
                        case 349:
                            seat.Score = 30;
                            break;
                        case 350:
                            seat.Score = 25;
                            break;
                        case 351:
                            seat.Score = 25;
                            break;
                        case 352:
                            seat.Score = 20;
                            break;
                        case 353:
                            seat.Score = 20;
                            break;
                        case 354:
                            seat.Score = 15;
                            break;
                        case 355:
                            seat.Score = 15;
                            break;
                        case 356:
                            seat.Score = 10;
                            break;
                        case 357:
                            seat.Score = 10;
                            break;
                        default:
                            seat.Score = 5;
                            break;
                    }
                }

                foreach (var seat in tables[3].Seats)
                {
                    switch (seat.SeatNumber)
                    {
                        case 400:
                            seat.Score = 40;
                            break;
                        case 401:
                            seat.Score = 40;
                            break;
                        case 402:
                            seat.Score = 45;
                            break;
                        case 403:
                            seat.Score = 45;
                            break;
                        case 404:
                            seat.Score = 50;
                            break;
                        case 405:
                            seat.Score = 50;
                            break;
                        case 406:
                            seat.Score = 55;
                            break;
                        case 407:
                            seat.Score = 55;
                            break;
                        case 408:
                            seat.Score = 60;
                            break;
                        case 409:
                            seat.Score = 60;
                            break;
                        case 410:
                            seat.Score = 65;
                            break;
                        case 411:
                            seat.Score = 65;
                            break;
                        case 412:
                            seat.Score = 70;
                            break;
                        case 413:
                            seat.Score = 70;
                            break;
                        case 414:
                            seat.Score = 75;
                            break;
                        case 415:
                            seat.Score = 75;
                            break;
                        case 416:
                            seat.Score = 70;
                            break;
                        case 417:
                            seat.Score = 70;
                            break;
                        case 418:
                            seat.Score = 65;
                            break;
                        case 419:
                            seat.Score = 65;
                            break;
                        case 420:
                            seat.Score = 60;
                            break;
                        case 421:
                            seat.Score = 60;
                            break;
                        case 422:
                            seat.Score = 55;
                            break;
                        case 423:
                            seat.Score = 55;
                            break;
                        case 424:
                            seat.Score = 50;
                            break;
                        case 425:
                            seat.Score = 50;
                            break;
                        case 426:
                            seat.Score = 45;
                            break;
                        case 427:
                            seat.Score = 45;
                            break;
                        case 428:
                            seat.Score = 40;
                            break;
                        case 429:
                            seat.Score = 40;
                            break;
                        case 430:
                            seat.Score = 35;
                            break;
                        case 431:
                            seat.Score = 35;
                            break;
                        case 432:
                            seat.Score = 30;
                            break;
                        case 433:
                            seat.Score = 30;
                            break;
                        case 434:
                            seat.Score = 25;
                            break;
                        case 435:
                            seat.Score = 25;
                            break;
                        case 436:
                            seat.Score = 20;
                            break;
                        case 437:
                            seat.Score = 20;
                            break;
                        case 438:
                            seat.Score = 15;
                            break;
                        case 439:
                            seat.Score = 15;
                            break;
                        case 440:
                            seat.Score = 10;
                            break;
                        case 441:
                            seat.Score = 10;
                            break;
                        case 442:
                            seat.Score = 10;
                            break;
                        default:
                            seat.Score = 5;
                            break;
                    }
                }

                foreach (var seat in tables[4].Seats)
                {
                    switch (seat.SeatNumber)
                    {
                        case 500:
                            seat.Score = 15;
                            break;
                        case 501:
                            seat.Score = 15;
                            break;
                        case 502:
                            seat.Score = 20;
                            break;
                        case 503:
                            seat.Score = 20;
                            break;
                        case 504:
                            seat.Score = 25;
                            break;
                        case 505:
                            seat.Score = 25;
                            break;
                        case 506:
                            seat.Score = 30;
                            break;
                        case 507:
                            seat.Score = 30;
                            break;
                        case 508:
                            seat.Score = 35;
                            break;
                        case 509:
                            seat.Score = 35;
                            break;
                        case 510:
                            seat.Score = 40;
                            break;
                        case 511:
                            seat.Score = 40;
                            break;
                        case 512:
                            seat.Score = 45;
                            break;
                        case 513:
                            seat.Score = 45;
                            break;
                        case 514:
                            seat.Score = 50;
                            break;
                        case 515:
                            seat.Score = 50;
                            break;
                        case 516:
                            seat.Score = 45;
                            break;
                        case 517:
                            seat.Score = 45;
                            break;
                        case 518:
                            seat.Score = 40;
                            break;
                        case 519:
                            seat.Score = 40;
                            break;
                        case 520:
                            seat.Score = 35;
                            break;
                        case 521:
                            seat.Score = 35;
                            break;
                        case 522:
                            seat.Score = 30;
                            break;
                        case 523:
                            seat.Score = 30;
                            break;
                        case 524:
                            seat.Score = 25;
                            break;
                        case 525:
                            seat.Score = 25;
                            break;
                        case 526:
                            seat.Score = 20;
                            break;
                        case 527:
                            seat.Score = 20;
                            break;
                        case 528:
                            seat.Score = 15;
                            break;
                        case 529:
                            seat.Score = 15;
                            break;
                        case 530:
                            seat.Score = 10;
                            break;
                        case 531:
                            seat.Score = 10;
                            break;
                        default:
                            seat.Score = 5;
                            break;
                    }
                }
            }
        }
    }
}
