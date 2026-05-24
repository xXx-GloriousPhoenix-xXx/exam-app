// Запустити: node generate_exam_answers_2.js
// Залежності: npm install docx
// Результат: exam_answers_2.docx

const {
    Document, Packer, Paragraph, TextRun, HeadingLevel,
    AlignmentType, LevelFormat, PageBreak, BorderStyle
  } = require('docx');
  const fs = require('fs');
  
  // ─────────────────────────────────────────────
  // ДОПОМІЖНІ ФУНКЦІЇ
  // ─────────────────────────────────────────────
  
  function heading1(text) {
    return new Paragraph({
      heading: HeadingLevel.HEADING_1,
      children: [new TextRun({ text, bold: true, size: 32 })],
      spacing: { before: 480, after: 200 },
      pageBreakBefore: true,
    });
  }
  
  function heading2(text) {
    return new Paragraph({
      heading: HeadingLevel.HEADING_2,
      children: [new TextRun({ text, bold: true, size: 26 })],
      spacing: { before: 300, after: 120 },
    });
  }
  
  function para(text) {
    return new Paragraph({
      children: [new TextRun({ text, size: 24 })],
      spacing: { before: 80, after: 80 },
      alignment: AlignmentType.JUSTIFIED,
    });
  }
  
  function bold(text) {
    return new Paragraph({
      children: [new TextRun({ text, bold: true, size: 24 })],
      spacing: { before: 120, after: 60 },
    });
  }
  
  function bullet(text) {
    return new Paragraph({
      numbering: { reference: 'bullets', level: 0 },
      children: [new TextRun({ text, size: 24 })],
      spacing: { before: 40, after: 40 },
    });
  }
  
  function code(text) {
    return new Paragraph({
      children: [new TextRun({ text, font: 'Courier New', size: 20 })],
      spacing: { before: 30, after: 30 },
      indent: { left: 720 },
    });
  }
  
  function separator() {
    return new Paragraph({
      children: [new TextRun({ text: '' })],
      border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: '2E75B6' } },
      spacing: { before: 200, after: 200 },
    });
  }
  
  // ─────────────────────────────────────────────
  // ЗМІСТ ВІДПОВІДЕЙ
  // ─────────────────────────────────────────────
  
  const content = [];
  
  // ══════════════════════════════════════════════
  // ПИТАННЯ 2
  // ══════════════════════════════════════════════
  content.push(heading1('Питання 2. Класифікація паралельних обчислювальних систем'));
  
  content.push(heading2('2.1. Таксономія Фліна (Flynn\'s Taxonomy)'));
  content.push(para(
    'Найвідоміша класифікація паралельних систем запропонована Майклом Фліном у 1966 році. ' +
    'Вона базується на двох незалежних вимірах: кількості потоків інструкцій (Instruction Streams) ' +
    'та кількості потоків даних (Data Streams), що обробляються одночасно.'
  ));
  content.push(bullet('SISD (Single Instruction, Single Data) — класичний однопроцесорний комп\'ютер фон Неймана. Одна інструкція обробляє один елемент даних за раз. Послідовне виконання.'));
  content.push(bullet('SIMD (Single Instruction, Multiple Data) — одна і та сама інструкція застосовується паралельно до множини елементів даних. Приклади: векторні процесори, GPU, набори інструкцій SSE/AVX у сучасних CPU. Ефективно для масових однотипних операцій (обробка зображень, фізичне моделювання).'));
  content.push(bullet('MISD (Multiple Instruction, Single Data) — кілька інструкцій застосовуються до одного потоку даних. Теоретична категорія; прикладом можуть слугувати відмовостійкі системи з N-кратним резервуванням, де різні вузли перевіряють одні й ті самі дані різними алгоритмами.'));
  content.push(bullet('MIMD (Multiple Instruction, Multiple Data) — кожен процесор виконує власний потік інструкцій над власними даними. Найбільш поширена категорія: багатоядерні CPU, кластери, суперкомп\'ютери. Підрозділяється на системи зі спільною пам\'яттю (SMP, NUMA) та розподіленою пам\'яттю (кластери).'));
  
  content.push(heading2('2.2. Класифікація за архітектурою пам\'яті'));
  content.push(bullet('UMA (Uniform Memory Access) — симетрична багатопроцесорна система (SMP): всі процесори мають однаковий час доступу до будь-якої ділянки спільної пам\'яті. Прості у програмуванні, але погано масштабуються (шина пам\'яті стає вузьким місцем).'));
  content.push(bullet('NUMA (Non-Uniform Memory Access) — кожен процесор (або група) має локальну пам\'ять з меншою затримкою; доступ до пам\'яті інших процесорів відбувається через міжпроцесорну шину і є повільнішим. Сучасні багатопроцесорні сервери — NUMA-архітектури. Вимагає NUMA-aware програмування.'));
  content.push(bullet('Distributed Memory — кожен вузол має власну приватну пам\'ять. Взаємодія виключно через передачу повідомлень (MPI). Ідеально масштабується до тисяч вузлів.'));
  content.push(bullet('Hybrid (SMP + кластер) — сучасні кластери складаються з вузлів із спільною пам\'яттю (NUMA/SMP), з\'єднаних мережею. Програмування: MPI між вузлами + OpenMP/Java threads всередині вузла.'));
  
  content.push(heading2('2.3. Класифікація за зв\'язністю'));
  content.push(bullet('Сильнозв\'язані системи: спільна пам\'ять, низька латентність (нс), висока пропускна здатність. Приклади: багатоядерні CPU, SMP-сервери.'));
  content.push(bullet('Слабкозв\'язані системи: окремі машини, з\'єднані мережею (мкс–мс латентність). Приклади: кластери, Grid, Cloud.'));
  content.push(separator());
  
  // ══════════════════════════════════════════════
  // ПИТАННЯ 4
  // ══════════════════════════════════════════════
  content.push(heading1('Питання 4. Системи із загальною та розподіленою пам\'яттю'));
  
  content.push(heading2('4.1. Системи зі спільною пам\'яттю (Shared Memory)'));
  content.push(para(
    'У системах зі спільною пам\'яттю всі процесори/ядра мають безпосередній доступ до ' +
    'єдиного адресного простору. Синхронізація доступу до спільних даних здійснюється через ' +
    'примітиви синхронізації: mutex, semaphore, monitor. '
  ));
  content.push(para(
    'Переваги: проста комунікація через спільні змінні, низька латентність, знайомі абстракції (Java threads). ' +
    'Недоліки: обмежена масштабованість (конфлікти за шину пам\'яті), складність синхронізації, ' +
    'race conditions, deadlocks. У Java вся JVM heap є спільною між потоками — звідси необхідність synchronized/volatile.'
  ));
  
  content.push(heading2('4.2. Проблема когерентності кешу'));
  content.push(para(
    'Кожне ядро має власний кеш (L1, L2). Якщо два ядра кешують одну й ту саму ділянку пам\'яті ' +
    'і одне з них змінює значення, виникає некогерентність. Протоколи когерентності кешу ' +
    '(MESI, MOESI) автоматично інвалідують застарілі копії. Java Memory Model (JMM) визначає ' +
    'правила видимості: happens-before гарантії для synchronized, volatile, final.'
  ));
  
  content.push(heading2('4.3. Системи з розподіленою пам\'яттю (Distributed Memory)'));
  content.push(para(
    'Кожен процес має власний приватний адресний простір. Процеси не можуть безпосередньо ' +
    'читати або писати в пам\'ять інших процесів. Єдиний спосіб взаємодії — явна передача ' +
    'повідомлень (message passing). Реалізується через MPI, socket API, RMI, gRPC.'
  ));
  content.push(para(
    'Переваги: відмінна масштабованість (тисячі вузлів), відсутність конфліктів за пам\'ять, ' +
    'природна відмовостійкість (процеси ізольовані). ' +
    'Недоліки: висока вартість комунікації, складність програмування, необхідність явної серіалізації даних.'
  ));
  
  content.push(heading2('4.4. Порівняльна таблиця'));
  content.push(bullet('Спільна пам\'ять: єдиний адресний простір, низька латентність (нс), синхронізація mutex/monitor, обмежена масштабованість, Java Threads'));
  content.push(bullet('Розподілена пам\'ять: приватна пам\'ять, висока латентність (мкс–мс), явна передача повідомлень, висока масштабованість, MPI/RMI'));
  content.push(bullet('Гібридна: SMP-вузли + мережа, найпоширеніша у сучасних HPC кластерах, MPI + OpenMP/Java threads'));
  content.push(separator());
  
  // ══════════════════════════════════════════════
  // ПИТАННЯ 6
  // ══════════════════════════════════════════════
  content.push(heading1('Питання 6. Багатопоточна технологія Java'));
  
  content.push(heading2('6.1. Потоки у JVM'));
  content.push(para(
    'Java з самого початку (JDK 1.0, 1995) мала вбудовану підтримку багатопоточності як частину ' +
    'мови і стандартної бібліотеки. Кожен Java-потік (java.lang.Thread) відображається на нативний ' +
    'потік ОС (1:1 модель на HotSpot JVM). JVM при запуску створює кілька службових потоків: ' +
    'main, GC threads, JIT compiler, finalizer, reference handler.'
  ));
  
  content.push(heading2('6.2. Java Memory Model (JMM)'));
  content.push(para(
    'JMM (JSR-133, Java 5+) визначає правила видимості змін між потоками. Без явної синхронізації ' +
    'компілятор, JIT і процесор мають право переставляти інструкції. Гарантії видимості надають:'
  ));
  content.push(bullet('synchronized блок/метод: всі записи перед виходом з synchronized видимі після входу іншого потоку в той самий монітор'));
  content.push(bullet('volatile поле: кожен запис у volatile видимий всім читачам без затримки'));
  content.push(bullet('final поля: після завершення конструктора значення final полів видимі всім потокам'));
  content.push(bullet('java.util.concurrent класи: надають happens-before гарантії (Semaphore.release() hb Semaphore.acquire())'));
  
  content.push(heading2('6.3. Засоби багатопоточності Java'));
  content.push(bold('Низький рівень (java.lang):'));
  content.push(bullet('Thread: клас для представлення потоку; start(), join(), interrupt(), sleep()'));
  content.push(bullet('Runnable: функціональний інтерфейс для задач без результату'));
  content.push(bullet('Object.wait() / notify() / notifyAll(): базова міжпотокова комунікація через монітор'));
  content.push(bullet('synchronized: ключове слово для захисту критичних секцій'));
  content.push(bullet('volatile: гарантія видимості змін змінної'));
  
  content.push(bold('Середній рівень (java.util.concurrent.locks):'));
  content.push(bullet('ReentrantLock: повторно вхідний lock з підтримкою tryLock, lockInterruptibly'));
  content.push(bullet('ReentrantReadWriteLock: розділяє читачів і письменників для підвищення паралелізму'));
  content.push(bullet('StampedLock (Java 8): оптимістичне читання для high-performance сценаріїв'));
  content.push(bullet('Condition: аналог wait/notify для явних локерів'));
  
  content.push(bold('Високий рівень (java.util.concurrent):'));
  content.push(bullet('ExecutorService / ThreadPoolExecutor: управління пулами потоків'));
  content.push(bullet('Future / CompletableFuture: асинхронні обчислення з ланцюжком операцій'));
  content.push(bullet('ForkJoinPool: work-stealing пул для рекурсивних задач'));
  content.push(bullet('Atomic-класи: lock-free атомарні операції (AtomicInteger, AtomicReference)'));
  content.push(bullet('Concurrent колекції: ConcurrentHashMap, CopyOnWriteArrayList, BlockingQueue'));
  content.push(bullet('Синхронізатори: CountDownLatch, CyclicBarrier, Semaphore, Phaser'));
  content.push(separator());
  
  // ══════════════════════════════════════════════
  // ПИТАННЯ 8
  // ══════════════════════════════════════════════
  content.push(heading1('Питання 8. Стани потоку та переходи між станами'));
  
  content.push(heading2('8.1. Перелік станів Thread.State'));
  content.push(para(
    'У Java стан потоку визначається перерахуванням Thread.State (Java 5+). ' +
    'Метод thread.getState() повертає поточний стан. Є 6 станів:'
  ));
  content.push(bullet('NEW — об\'єкт Thread створено, але start() ще не викликано. Потік не існує в ОС.'));
  content.push(bullet('RUNNABLE — потік виконується або готовий до виконання (очікує на виділення CPU планувальником ОС). Включає два підстани: "running" і "ready".'));
  content.push(bullet('BLOCKED — потік намагається увійти в synchronized блок або метод, монітор якого утримується іншим потоком. Чекає звільнення монітора.'));
  content.push(bullet('WAITING — потік чекає невизначено довго. Причини: Object.wait() без таймауту, Thread.join() без таймауту, LockSupport.park().'));
  content.push(bullet('TIMED_WAITING — потік чекає обмежений час. Причини: Thread.sleep(ms), Object.wait(ms), Thread.join(ms), LockSupport.parkNanos().'));
  content.push(bullet('TERMINATED — виконання методу run() завершено (нормально або через виняток). Потік не може бути перезапущений.'));
  
  content.push(heading2('8.2. Діаграма переходів між станами'));
  content.push(bullet('NEW → RUNNABLE: виклик thread.start()'));
  content.push(bullet('RUNNABLE → BLOCKED: спроба увійти в зайнятий synchronized блок'));
  content.push(bullet('BLOCKED → RUNNABLE: монітор звільнено, потік може його захопити'));
  content.push(bullet('RUNNABLE → WAITING: виклик wait(), join(), park()'));
  content.push(bullet('WAITING → RUNNABLE: виклик notify()/notifyAll(), завершення join(), unpark()'));
  content.push(bullet('RUNNABLE → TIMED_WAITING: виклик sleep(ms), wait(ms), join(ms)'));
  content.push(bullet('TIMED_WAITING → RUNNABLE: закінчився таймаут або прийшов notify()'));
  content.push(bullet('RUNNABLE → TERMINATED: завершення run()'));
  
  content.push(heading2('8.3. Практичне значення'));
  content.push(para(
    'Знання станів критично для діагностики проблем: якщо потік застряг у BLOCKED — ' +
    'ймовірний deadlock або lock contention; якщо у WAITING — можливо не надійшов notify(); ' +
    'якщо більшість потоків TIMED_WAITING — система гальмує через надмірний sleep(). ' +
    'Thread dumps (jstack, VisualVM) показують стан кожного потоку і стек викликів, ' +
    'що є головним інструментом діагностики багатопоточних проблем.'
  ));
  content.push(separator());
  
  // ══════════════════════════════════════════════
  // ПИТАННЯ 10
  // ══════════════════════════════════════════════
  content.push(heading1('Питання 10. Алгоритми паралельного множення матриць'));
  
  content.push(heading2('10.1. Огляд підходів'));
  content.push(para(
    'Множення матриць C = A × B є еталонним завданням для паралельних обчислень. ' +
    'Елемент c[i][j] = Σ a[i][k] * b[k][j]. Складність: O(n³) для наївного алгоритму. ' +
    'Для великих матриць (n ≥ 512) паралелізація дає суттєвий виграш. ' +
    'Різні алгоритми оптимізовані під різні архітектури.'
  ));
  
  content.push(heading2('10.2. Стрічковий алгоритм (Row Strip Decomposition)'));
  content.push(para(
    'Матриця A ділиться на горизонтальні смуги між p потоками. Потік i обробляє рядки ' +
    'від i*(n/p) до (i+1)*(n/p). Матриця B — shared read-only. Реалізується через ' +
    'FixedThreadPool або ForkJoinPool у Java. Дуже простий у реалізації, ефективний ' +
    'для shared-memory систем. Ефективність наближається до 100% при правильному розмірі задачі.'
  ));
  
  content.push(heading2('10.3. Стовпчиковий алгоритм (Column Strip Decomposition)'));
  content.push(para(
    'Симетричний варіант: матриця B ділиться на вертикальні смуги. Потік i обчислює ' +
    'стовпці результуючої матриці C від i*(m/p) до (i+1)*(m/p). Матриця A — shared read-only. ' +
    'Однаковий баланс навантаження, але гірша локальність кешу при доступі до стовпців B (column-major access).'
  ));
  
  content.push(heading2('10.4. Блоковий алгоритм (Block Decomposition)'));
  content.push(para(
    'Матриці A, B і C ділять на блоки розміру (n/√p)×(n/√p). Кожен процес відповідає за ' +
    'блок результату C[i][j] і обчислює його як суму добутків блоків A[i][k] × B[k][j] по всіх k. ' +
    'Краща локальність кешу, менший обсяг комунікацій у розподілених системах. ' +
    'Основа алгоритмів Кеннона і Фокса.'
  ));
  
  content.push(heading2('10.5. Алгоритм Штрассена (Strassen)'));
  content.push(para(
    'Зменшує кількість множень з 8 до 7 для 2×2 блоків, що дає складність O(n^2.81) замість O(n³). ' +
    'Погано паралелізується через нерегулярну структуру залежностей, але може комбінуватись ' +
    'з блоковою декомпозицією на верхніх рівнях рекурсії.'
  ));
  content.push(separator());
  
  // ══════════════════════════════════════════════
  // ПИТАННЯ 12
  // ══════════════════════════════════════════════
  content.push(heading1('Питання 12. Алгоритм Фокса паралельного множення матриць'));
  
  content.push(heading2('12.1. Постановка задачі та ідея'));
  content.push(para(
    'Алгоритм Фокса (Fox\'s Algorithm, також SUMMA — Scalable Universal Matrix Multiply Algorithm) — ' +
    'блоковий алгоритм паралельного множення матриць для 2D-сітки процесів p×p. ' +
    'На відміну від алгоритму Кеннона, замість циклічних зсувів використовує broadcast ' +
    '(розсилку) блоків матриці A по рядках сітки процесів.'
  ));
  
  content.push(heading2('12.2. Розподіл даних'));
  content.push(para(
    'Матриці A, B і C розбиваються на блоки розміру (n/p)×(n/p). ' +
    'Процес (i, j) в 2D-сітці зберігає блоки A_ij, B_ij і акумулює C_ij. ' +
    'Для правильної роботи матриця має бути квадратною, а p — точним квадратом.'
  ));
  
  content.push(heading2('12.3. Покроковий опис алгоритму'));
  content.push(bullet('Ініціалізація: C_ij = 0 для всіх (i, j)'));
  content.push(bullet('Цикл по k від 0 до p-1:'));
  content.push(bullet('  Broadcast: процес (i, (i+k) mod p) розсилає блок A_i,(i+k)mod p всім процесам рядка i (MPI_Bcast по рядках сітки)'));
  content.push(bullet('  Локальне множення: кожен процес (i,j) обчислює C_ij += A_broadcast * B_ij'));
  content.push(bullet('  Зсув B: блоки матриці B зсуваються вгору на 1 позицію по стовпцях сітки (MPI_Sendrecv)'));
  
  content.push(heading2('12.4. Порівняння з алгоритмом Кеннона'));
  content.push(bullet('Кеннон: початковий зсув + p кроків з shift; рівномірна комунікація, ідеально для 2D-торусної мережі'));
  content.push(bullet('Фокс/SUMMA: broadcast + shift; гнучкіший (не потребує квадратної сітки у розширених версіях), краще для нерегулярних мереж'));
  content.push(bullet('Обидва: O(n³/p) обчислень, O(n²/√p) комунікацій, ефективність ~90%+ на кластерах'));
  
  content.push(heading2('12.5. Java реалізація (концептуально)'));
  content.push(code('// p^0.5 x p^0.5 сітка потоків, кожен потік — "процес" сітки'));
  content.push(code('for (int step = 0; step < sqrtP; step++) {'));
  content.push(code('    // 1. Broadcast блоку A по рядку'));
  content.push(code('    int broadcastCol = (row + step) % sqrtP;'));
  content.push(code('    double[][] broadcastBlock = (col == broadcastCol)'));
  content.push(code('        ? localA : receiveFromRowBroadcast(row, broadcastCol);'));
  content.push(code('    // 2. Локальне множення'));
  content.push(code('    multiplyAndAdd(localC, broadcastBlock, localB);'));
  content.push(code('    // 3. Циклічний зсув B вгору'));
  content.push(code('    localB = shiftBUp(localB, row, col, sqrtP);'));
  content.push(code('}'));
  content.push(separator());
  
  // ══════════════════════════════════════════════
  // ПИТАННЯ 14
  // ══════════════════════════════════════════════
  content.push(heading1('Питання 14. Призупинка та відновлення роботи потоку. Переривання дії потоку'));
  
  content.push(heading2('14.1. Призупинка потоку: Thread.sleep()'));
  content.push(para(
    'Thread.sleep(long millis) призупиняє поточний потік на вказану кількість мілісекунд, ' +
    'переводячи його в стан TIMED_WAITING. Потік не звільняє утримувані монітори (на відміну від wait()). ' +
    'Може бути перерваний через interrupt() — тоді кидається InterruptedException.'
  ));
  content.push(code('try {'));
  content.push(code('    Thread.sleep(1000); // пауза 1 секунда'));
  content.push(code('} catch (InterruptedException e) {'));
  content.push(code('    Thread.currentThread().interrupt(); // відновити прапор переривання'));
  content.push(code('    return; // коректно завершити роботу'));
  content.push(code('}'));
  
  content.push(heading2('14.2. Очікування завершення: join()'));
  content.push(para(
    'thread.join() призупиняє поточний потік у стані WAITING до завершення потоку thread. ' +
    'thread.join(long timeout) — з таймаутом, не довше timeout мс (стан TIMED_WAITING). ' +
    'Типовий паттерн: запустити N worker-потоків, потім join() кожного для збору результатів.'
  ));
  content.push(code('List<Thread> workers = new ArrayList<>();'));
  content.push(code('for (int i = 0; i < 4; i++) {'));
  content.push(code('    Thread t = new Thread(new ComputeTask(i));'));
  content.push(code('    workers.add(t);'));
  content.push(code('    t.start();'));
  content.push(code('}'));
  content.push(code('for (Thread t : workers) {'));
  content.push(code('    t.join(); // чекаємо кожного'));
  content.push(code('}'));
  content.push(code('// тут всі потоки завершені, результати готові'));
  
  content.push(heading2('14.3. Переривання потоку: interrupt()'));
  content.push(para(
    'Java не має механізму примусового зупинення потоку (Thread.stop() застарілий і небезпечний: ' +
    'може залишити об\'єкти в неконсистентному стані). Замість цього — кооперативний підхід: ' +
    'виклик t.interrupt() встановлює внутрішній прапор переривання. Два сценарії:'
  ));
  content.push(bullet('Потік заблокований (sleep, wait, join): негайно кидається InterruptedException, прапор очищається. Обов\'язково обробляти!'));
  content.push(bullet('Потік виконується: прапор встановлюється, потік сам має перевіряти Thread.currentThread().isInterrupted() і коректно завершуватись.'));
  
  content.push(code('// Правильний паттерн для потоку, що підтримує переривання:'));
  content.push(code('public void run() {'));
  content.push(code('    while (!Thread.currentThread().isInterrupted()) {'));
  content.push(code('        try {'));
  content.push(code('            processNextItem();'));
  content.push(code('            Thread.sleep(100);'));
  content.push(code('        } catch (InterruptedException e) {'));
  content.push(code('            Thread.currentThread().interrupt(); // відновити прапор'));
  content.push(code('            break; // вийти з циклу'));
  content.push(code('        }'));
  content.push(code('    }'));
  content.push(code('    cleanup(); // звільнити ресурси перед виходом'));
  content.push(code('}'));
  
  content.push(heading2('14.4. LockSupport.park() / unpark()'));
  content.push(para(
    'Низькорівневий механізм призупинки/відновлення. park() призупиняє поточний потік (WAITING/TIMED_WAITING). ' +
    'unpark(thread) дозволяє відновити конкретний потік. Використовується всередині ReentrantLock, ' +
    'CompletableFuture та інших concurrent примітивів. На відміну від wait/notify, unpark можна ' +
    'викликати до park() — потік не заблокується.'
  ));
  content.push(separator());
  
  // ══════════════════════════════════════════════
  // ПИТАННЯ 16
  // ══════════════════════════════════════════════
  content.push(heading1('Питання 16. Синхронізація в паралельних обчисленнях'));
  
  content.push(heading2('16.1. Поняття race condition'));
  content.push(para(
    'Race condition (перегонний стан) — помилка, при якій результат програми залежить від ' +
    'відносного порядку виконання потоків. Класичний приклад: два потоки інкрементують ' +
    'спільний лічильник. Операція count++ насправді складається з трьох: read, increment, write. ' +
    'Якщо обидва потоки прочитають одне і те саме значення і обидва запишуть +1, ' +
    'результат буде count+1 замість count+2. Синхронізація усуває цю проблему.'
  ));
  
  content.push(heading2('16.2. Механізми синхронізації в Java'));
  content.push(bold('Монітор (synchronized):'));
  content.push(code('synchronized (sharedObject) {'));
  content.push(code('    // критична секція — виконується атомарно'));
  content.push(code('    sharedCounter++;'));
  content.push(code('}'));
  
  content.push(bold('Явний Lock (ReentrantLock):'));
  content.push(code('lock.lock();'));
  content.push(code('try { sharedCounter++; }'));
  content.push(code('finally { lock.unlock(); }'));
  
  content.push(bold('Атомарні змінні (без блокування):'));
  content.push(code('AtomicInteger counter = new AtomicInteger(0);'));
  content.push(code('counter.incrementAndGet(); // атомарно, lock-free'));
  
  content.push(heading2('16.3. Синхронізаційні бар\'єри'));
  content.push(para(
    'Бар\'єрна синхронізація — усі потоки/процеси чекають один одного у певній точці ' +
    'і продовжують роботу лише після того, як усі досягли бар\'єра.'
  ));
  content.push(code('CyclicBarrier barrier = new CyclicBarrier(numThreads, () -> {'));
  content.push(code('    // виконується одним потоком після досягнення бар\'єра всіма'));
  content.push(code('    mergeResults();'));
  content.push(code('});'));
  content.push(code('// У кожному потоці:'));
  content.push(code('computePartialResult();'));
  content.push(code('barrier.await(); // чекати інших'));
  content.push(code('// продовжити після бар\'єра'));
  
  content.push(heading2('16.4. Семафори'));
  content.push(para(
    'Semaphore — лічильний семафор, що обмежує кількість одночасних доступів до ресурсу. ' +
    'acquire() зменшує лічильник (блокує при 0), release() збільшує. ' +
    'Бінарний семафор (permits=1) — аналог mutex. ' +
    'Корисний для обмеження кількості одночасних підключень до БД, API rate limiting тощо.'
  ));
  content.push(code('Semaphore semaphore = new Semaphore(3); // 3 одночасних доступи'));
  content.push(code('semaphore.acquire();'));
  content.push(code('try { accessDatabase(); }'));
  content.push(code('finally { semaphore.release(); }'));
  content.push(separator());
  
  // ══════════════════════════════════════════════
  // ПИТАННЯ 18
  // ══════════════════════════════════════════════
  content.push(heading1('Питання 18. Блокування об\'єкту'));
  
  content.push(heading2('18.1. Монітор об\'єкта в Java'));
  content.push(para(
    'Кожен об\'єкт у Java має вбудований монітор (intrinsic lock, або monitor lock). ' +
    'Це прихований mutex, що автоматично захоплюється при вході в synchronized блок/метод ' +
    'і звільняється при виході. Монітор є реєнтерабельним (reentrant): один і той самий потік ' +
    'може входити в synchronized кілька разів (лічильник входжень збільшується).'
  ));
  
  content.push(heading2('18.2. Synchronized блоки та методи'));
  content.push(code('// Блокування на рівні екземпляра (монітор = this):'));
  content.push(code('public synchronized void increment() { count++; }'));
  content.push(code(''));
  content.push(code('// Еквівалентно:'));
  content.push(code('public void increment() {'));
  content.push(code('    synchronized (this) { count++; }'));
  content.push(code('}'));
  content.push(code(''));
  content.push(code('// Блокування на рівні класу (монітор = MyClass.class):'));
  content.push(code('public static synchronized void staticMethod() { ... }'));
  content.push(code(''));
  content.push(code('// Блокування на окремому об\'єкті-локері (краща практика):'));
  content.push(code('private final Object lock = new Object();'));
  content.push(code('public void safeMethod() {'));
  content.push(code('    synchronized (lock) { ... }'));
  content.push(code('}'));
  
  content.push(heading2('18.3. Реєнтерабельність (Reentrancy)'));
  content.push(code('class ReentrantExample {'));
  content.push(code('    public synchronized void outer() {'));
  content.push(code('        inner(); // OK: той самий потік вже тримає монітор'));
  content.push(code('    }'));
  content.push(code('    public synchronized void inner() {'));
  content.push(code('        // лічильник входжень = 2, не deadlock'));
  content.push(code('    }'));
  content.push(code('}'));
  
  content.push(heading2('18.4. Volatile vs Synchronized'));
  content.push(para(
    'volatile гарантує лише видимість (кожен запис видимий всім читачам), але не атомарність. ' +
    'Підходить для простих прапорців (boolean running) і одиночних reads/writes. ' +
    'synchronized гарантує і видимість, і атомарність критичної секції. ' +
    'Потрібен для compound operations (check-then-act, read-modify-write).'
  ));
  
  content.push(heading2('18.5. ReadWriteLock'));
  content.push(para(
    'ReentrantReadWriteLock дозволяє кільком потокам одночасно читати, але забезпечує ' +
    'ексклюзивний доступ при записі. Підходить для структур, що читаються часто, але рідко пишуться.'
  ));
  content.push(code('ReadWriteLock rwLock = new ReentrantReadWriteLock();'));
  content.push(code('// Читання (кілька потоків одночасно):'));
  content.push(code('rwLock.readLock().lock();'));
  content.push(code('try { return data.get(key); }'));
  content.push(code('finally { rwLock.readLock().unlock(); }'));
  content.push(code('// Запис (ексклюзивно):'));
  content.push(code('rwLock.writeLock().lock();'));
  content.push(code('try { data.put(key, value); }'));
  content.push(code('finally { rwLock.writeLock().unlock(); }'));
  content.push(separator());
  
  // ══════════════════════════════════════════════
  // ПИТАННЯ 20
  // ══════════════════════════════════════════════
  content.push(heading1('Питання 20. Проблеми управління потоками: дедлок та інші небажані стани'));
  
  content.push(heading2('20.1. Дедлок (Deadlock)'));
  content.push(para(
    'Дедлок — ситуація, де два або більше потоки назавжди заблоковані в очікуванні ' +
    'ресурсів, які утримують один одного. Умови Коффмана (всі 4 потрібні для дедлоку): ' +
    '(1) взаємне виключення, (2) утримування і очікування, (3) відсутність примусового звільнення, ' +
    '(4) циклічне очікування.'
  ));
  content.push(code('// Класичний deadlock:'));
  content.push(code('// Потік 1: lock(A), потім lock(B)'));
  content.push(code('// Потік 2: lock(B), потім lock(A)  ← циклічне очікування'));
  content.push(para(
    'Запобігання: завжди захоплювати локери в одному глобальному порядку; ' +
    'використовувати tryLock(timeout) замість lock(); уникати вкладених synchronized блоків; ' +
    'використовувати структури даних без блокувань (lock-free).'
  ));
  
  content.push(heading2('20.2. Livelock'));
  content.push(para(
    'Livelock — потоки активні (не заблоковані), але не досягають прогресу, бо постійно ' +
    'реагують один на одного. Аналогія: два люди в коридорі, що нескінченно поступаються місцем. ' +
    'Вирішення: ввести випадкову затримку або пріоритети при конфлікті.'
  ));
  
  content.push(heading2('20.3. Starvation (голодування)'));
  content.push(para(
    'Starvation — потік ніколи не отримує необхідний ресурс, бо інші постійно мають вищий пріоритет ' +
    'або "удачу" при конкуренції. Вирішення: fair lock (ReentrantLock(true) — fair mode), ' +
    'рівні пріоритети, обмежене очікування.'
  ));
  
  content.push(heading2('20.4. Race Condition'));
  content.push(para(
    'Результат залежить від порядку планування потоків. Виявляється рідко при тестуванні, ' +
    'але проявляється у виробництві під навантаженням. Інструменти виявлення: ThreadSanitizer, ' +
    'Java PathFinder, FindBugs/SpotBugs (@GuardedBy анотація).'
  ));
  
  content.push(heading2('20.5. Діагностика'));
  content.push(bullet('jstack <pid>: thread dump з поточним станом і стеком кожного потоку'));
  content.push(bullet('VisualVM / JConsole: графічний моніторинг потоків, виявлення BLOCKED стану'));
  content.push(bullet('Java Flight Recorder: низькорівневий трасувальник подій JVM'));
  content.push(bullet('Deadlock detection: JVM автоматично виявляє deadlocks між synchronized потоками і виводить у thread dump'));
  content.push(separator());
  
  // ══════════════════════════════════════════════
  // ПИТАННЯ 22
  // ══════════════════════════════════════════════
  content.push(heading1('Питання 22. Локери та управління потоками'));
  
  content.push(heading2('22.1. Недоліки synchronized та потреба у явних локерах'));
  content.push(para(
    'Вбудований synchronized має обмеження: неможливо перервати очікування на монітор; ' +
    'неможливо спробувати захопити з таймаутом; один монітор — одна умова очікування; ' +
    'неможливо справедливо розподілити доступ між потоками (fairness). ' +
    'java.util.concurrent.locks вирішує всі ці проблеми.'
  ));
  
  content.push(heading2('22.2. ReentrantLock'));
  content.push(code('ReentrantLock lock = new ReentrantLock(true); // fair=true'));
  content.push(code(''));
  content.push(code('// Звичайне блокування:'));
  content.push(code('lock.lock();'));
  content.push(code('try { /* критична секція */ }'));
  content.push(code('finally { lock.unlock(); } // обов\'язково у finally!'));
  content.push(code(''));
  content.push(code('// Спроба з таймаутом:'));
  content.push(code('if (lock.tryLock(500, TimeUnit.MILLISECONDS)) {'));
  content.push(code('    try { /* робота */ }'));
  content.push(code('    finally { lock.unlock(); }'));
  content.push(code('} else {'));
  content.push(code('    // не вдалось захопити — обробити ситуацію'));
  content.push(code('}'));
  content.push(code(''));
  content.push(code('// Переривне очікування:'));
  content.push(code('lock.lockInterruptibly(); // кидає InterruptedException при interrupt()'));
  
  content.push(heading2('22.3. Condition — кілька умов для одного Lock'));
  content.push(code('ReentrantLock lock = new ReentrantLock();'));
  content.push(code('Condition notFull  = lock.newCondition();'));
  content.push(code('Condition notEmpty = lock.newCondition();'));
  content.push(code(''));
  content.push(code('// Producer:'));
  content.push(code('lock.lock();'));
  content.push(code('try {'));
  content.push(code('    while (isFull()) notFull.await();'));
  content.push(code('    buffer.add(item);'));
  content.push(code('    notEmpty.signal(); // розбудити одного споживача'));
  content.push(code('} finally { lock.unlock(); }'));
  content.push(code(''));
  content.push(code('// Consumer:'));
  content.push(code('lock.lock();'));
  content.push(code('try {'));
  content.push(code('    while (isEmpty()) notEmpty.await();'));
  content.push(code('    item = buffer.remove();'));
  content.push(code('    notFull.signal();'));
  content.push(code('} finally { lock.unlock(); }'));
  
  content.push(heading2('22.4. StampedLock (Java 8)'));
  content.push(para(
    'StampedLock — оптимізований локер для сценаріїв з переважним читанням. ' +
    'Підтримує три режими: writing (ексклюзивний), reading (спільний), optimisticRead. ' +
    'Optimistic read не блокує, але потребує validate() для перевірки: якщо між читанням ' +
    'і validate() відбувся запис, потрібно повторити з read lock. ' +
    'Продуктивніший за ReadWriteLock при низькій конкуренції на запис.'
  ));
  content.push(separator());
  
  // ══════════════════════════════════════════════
  // ПИТАННЯ 24
  // ══════════════════════════════════════════════
  content.push(heading1('Питання 24. Інтерфейс Executor та бібліотечні класи'));
  
  content.push(heading2('24.1. Ієрархія інтерфейсів'));
  content.push(bullet('Executor: базовий інтерфейс, один метод execute(Runnable). Відокремлює подачу задачі від механізму виконання.'));
  content.push(bullet('ExecutorService extends Executor: додає submit(Callable), shutdown(), awaitTermination(), invokeAll(), invokeAny()'));
  content.push(bullet('ScheduledExecutorService extends ExecutorService: додає schedule(), scheduleAtFixedRate(), scheduleWithFixedDelay()'));
  content.push(bullet('AbstractExecutorService: базова реалізація з newTaskFor()'));
  content.push(bullet('ThreadPoolExecutor extends AbstractExecutorService: основна реалізація пулів потоків'));
  content.push(bullet('ScheduledThreadPoolExecutor: реалізація планувальника'));
  content.push(bullet('ForkJoinPool extends AbstractExecutorService: work-stealing пул'));
  
  content.push(heading2('24.2. Фабричний клас Executors'));
  content.push(code('// Фіксований пул (CPU-bound задачі):'));
  content.push(code('ExecutorService fixed = Executors.newFixedThreadPool(n);'));
  content.push(code(''));
  content.push(code('// Динамічний пул (I/O-bound задачі):'));
  content.push(code('ExecutorService cached = Executors.newCachedThreadPool();'));
  content.push(code(''));
  content.push(code('// Один потік (послідовна обробка черги):'));
  content.push(code('ExecutorService single = Executors.newSingleThreadExecutor();'));
  content.push(code(''));
  content.push(code('// Планувальник:'));
  content.push(code('ScheduledExecutorService sched = Executors.newScheduledThreadPool(2);'));
  content.push(code(''));
  content.push(code('// Work-stealing (ForkJoin під капотом, Java 8):'));
  content.push(code('ExecutorService ws = Executors.newWorkStealingPool();'));
  
  content.push(heading2('24.3. Lifecycle ExecutorService'));
  content.push(bullet('Running: приймає і виконує задачі (початковий стан)'));
  content.push(bullet('Shutting down: shutdown() — нові задачі не приймаються, поточні доробляються'));
  content.push(bullet('Terminated: всі задачі завершені після shutdown()'));
  content.push(bullet('shutdownNow(): намагається перервати виконувані задачі (interrupt), повертає список незапущених'));
  
  content.push(heading2('24.4. invokeAll та invokeAny'));
  content.push(code('List<Callable<Integer>> tasks = List.of(task1, task2, task3);'));
  content.push(code(''));
  content.push(code('// invokeAll: чекає завершення ВСІХ задач'));
  content.push(code('List<Future<Integer>> futures = executor.invokeAll(tasks);'));
  content.push(code('for (Future<Integer> f : futures) {'));
  content.push(code('    System.out.println(f.get());'));
  content.push(code('}'));
  content.push(code(''));
  content.push(code('// invokeAny: повертає результат ПЕРШОЇ успішно завершеної задачі'));
  content.push(code('Integer firstResult = executor.invokeAny(tasks);'));
  content.push(separator());
  
  // ══════════════════════════════════════════════
  // ПИТАННЯ 26
  // ══════════════════════════════════════════════
  content.push(heading1('Питання 26. Методи моделювання паралельних обчислень'));
  
  content.push(heading2('26.1. Навіщо моделювати'));
  content.push(para(
    'Паралельні програми надзвичайно складно тестувати: поведінка залежить від планувальника ОС, ' +
    'навантаження, апаратури. Деякі баги відтворюються один раз на мільйон запусків. ' +
    'Формальне моделювання дозволяє: аналізувати коректність до написання коду, ' +
    'виявляти deadlocks і race conditions, оцінювати теоретичну продуктивність.'
  ));
  
  content.push(heading2('26.2. Граф потоку управління (Control Flow Graph)'));
  content.push(para(
    'Представляє програму як граф станів і переходів. Паралельна програма — декілька таких ' +
    'графів, що виконуються одночасно. Для аналізу deadlocks будується граф ресурсів (Resource Allocation Graph): ' +
    'вузли — потоки і ресурси, дуги — запити і призначення. Цикл у графі → потенційний deadlock.'
  ));
  
  content.push(heading2('26.3. Модель акторів (Actor Model)'));
  content.push(para(
    'Кожен актор — незалежна одиниця обчислення зі своїм станом і поштовою скринькою. ' +
    'Актори взаємодіють лише через асинхронні повідомлення, не поділяють стану. ' +
    'Відсутні race conditions за визначенням. У Java: Akka framework (Java/Scala). ' +
    'Ця модель природно підходить для розподілених систем.'
  ));
  
  content.push(heading2('26.4. CSP (Communicating Sequential Processes)'));
  content.push(para(
    'Формальна модель (Хоар, 1978): паралельні процеси спілкуються через синхронні канали. ' +
    'Процеси не мають спільного стану. В Java часткова реалізація через SynchronousQueue. ' +
    'У Go мові канали є основним механізмом — ідея безпосередньо з CSP.'
  ));
  
  content.push(heading2('26.5. Мережі Петрі'));
  content.push(para(
    'Графічна математична модель для опису паралельних систем. Дозволяє моделювати ' +
    'синхронізацію, конкуренцію за ресурси, deadlock. Детально — наступне питання.'
  ));
  content.push(separator());
  
  // ══════════════════════════════════════════════
  // ПИТАННЯ 28
  // ══════════════════════════════════════════════
  content.push(heading1('Питання 28. Моделювання паралельних програм мережею Петрі'));
  
  content.push(heading2('28.1. Структура мережі Петрі'));
  content.push(para(
    'Мережа Петрі (Petri Net, 1962) — орієнтований двочастковий граф: ' +
    'місця (places, кола) представляють стани або ресурси; ' +
    'переходи (transitions, прямокутники) представляють події або дії; ' +
    "дуги з'єднують місця з переходами і навпаки; " +
    'мітки (tokens, жетони) у місцях відображають поточний стан системи.'
  ));
  
  content.push(heading2('28.2. Правила спрацьовування'));
  content.push(para(
    'Перехід є дозволеним (enabled), якщо кожне його вхідне місце містить хоча б стільки жетонів, ' +
    'скільки вимагає вхідна дуга. При спрацьовуванні (firing): з кожного вхідного місця ' +
    'забирається відповідна кількість жетонів, у кожне вихідне місце додається відповідна кількість. ' +
    'Декілька переходів можуть бути дозволеними одночасно — моделює паралелізм.'
  ));
  
  content.push(heading2('28.3. Моделювання синхронізаційних примітивів'));
  content.push(bullet('Mutex: місце з одним жетоном — "ресурс вільний". Перехід "захопити" забирає жетон, перехід "звільнити" повертає.'));
  content.push(bullet('Семафор: місце з N жетонами. acquire забирає 1, release додає 1.'));
  content.push(bullet('Producer-Consumer: місце "буфер" обмежене N жетонами, producer додає, consumer забирає.'));
  content.push(bullet('Deadlock: розмітка, з якої жоден перехід не є дозволеним — тупик у мережі Петрі.'));
  
  content.push(heading2('28.4. Аналіз мереж Петрі'));
  content.push(bullet('Досяжність (Reachability): чи можна досягти певної розмітки? Виявляє deadlock.'));
  content.push(bullet('Обмеженість (Boundedness): чи обмежена кількість жетонів у кожному місці? Безмежна мережа → переповнення ресурсу.'));
  content.push(bullet('Живість (Liveness): чи може кожен перехід рано чи пізно спрацювати? Мертвий перехід → недосяжний код або starvation.'));
  content.push(bullet('Інваріанти: математичні умови, що зберігаються при будь-якій розмітці.'));
  
  content.push(heading2('28.5. Інструменти'));
  content.push(para(
    'PIPE (Platform Independent Petri Net Editor), GreatSPN, TINA — безкоштовні інструменти ' +
    'для візуального створення і аналізу мереж Петрі. Дозволяють будувати граф досяжності ' +
    'і автоматично виявляти deadlocks і liveness-порушення.'
  ));
  content.push(separator());
  
  // ══════════════════════════════════════════════
  // ПИТАННЯ 30
  // ══════════════════════════════════════════════
  content.push(heading1('Питання 30. Алгоритми паралельного сумування та оцінка ефективності'));
  
  content.push(heading2('30.1. Паралельна редукція (Parallel Reduction)'));
  content.push(para(
    'Задача: обчислити суму N елементів масиву паралельно. Послідовний алгоритм: O(N) операцій, O(N) часу. ' +
    'Паралельна редукція: дерево сумувань. На кожному рівні сусідні пари елементів сумуються паралельно. ' +
    'Глибина дерева: ⌈log₂(N)⌉. При p = N/2 процесорів: час T(N, N/2) = O(log N).'
  ));
  
  content.push(heading2('30.2. Реалізація в Java (ForkJoin)'));
  content.push(code('class SumTask extends RecursiveTask<Long> {'));
  content.push(code('    private final long[] data;'));
  content.push(code('    private final int lo, hi;'));
  content.push(code('    static final int THRESHOLD = 1000;'));
  content.push(code(''));
  content.push(code('    protected Long compute() {'));
  content.push(code('        if (hi - lo <= THRESHOLD) {'));
  content.push(code('            long sum = 0;'));
  content.push(code('            for (int i = lo; i < hi; i++) sum += data[i];'));
  content.push(code('            return sum;'));
  content.push(code('        }'));
  content.push(code('        int mid = (lo + hi) >>> 1;'));
  content.push(code('        SumTask left  = new SumTask(data, lo, mid);'));
  content.push(code('        SumTask right = new SumTask(data, mid, hi);'));
  content.push(code('        left.fork();'));
  content.push(code('        return right.compute() + left.join();'));
  content.push(code('    }'));
  content.push(code('}'));
  content.push(code('long total = ForkJoinPool.commonPool()'));
  content.push(code('              .invoke(new SumTask(data, 0, data.length));'));
  
  content.push(heading2('30.3. Оцінка ефективності'));
  content.push(para(
    'T₁ = N-1 операцій (послідовне сумування). ' +
    'T∞ = ⌈log₂(N)⌉ операцій (критичний шлях дерева). ' +
    'Ступінь паралелізму = (N-1)/⌈log₂(N)⌉ ≈ N/log₂(N). ' +
    'Для p << N/log(N): прискорення ≈ p (ідеальне). ' +
    'При p = N/2: прискорення ≈ N/log(N). Ефективність E = N/(N/2 * log(N)) = 2/log(N) — ' +
    'знижується при збільшенні N через незайнятість процесорів на верхніх рівнях дерева.'
  ));
  
  content.push(heading2('30.4. Паралельне сумування у MPI'));
  content.push(code('// Кожен процес рахує часткову суму:'));
  content.push(code('double localSum = 0;'));
  content.push(code('for (int i = rank * chunk; i < (rank+1) * chunk; i++)'));
  content.push(code('    localSum += array[i];'));
  content.push(code(''));
  content.push(code('// MPI_Reduce: дерево редукції реалізовано автоматично'));
  content.push(code('double[] globalSum = {0};'));
  content.push(code('MPI.COMM_WORLD.Reduce(new double[]{localSum}, 0,'));
  content.push(code('                     globalSum, 0, 1, MPI.DOUBLE, MPI.SUM, 0);'));
  content.push(separator());
  
  // ══════════════════════════════════════════════
  // ПИТАННЯ 32
  // ══════════════════════════════════════════════
  content.push(heading1('Питання 32. Експериментальне дослідження ефективності паралельних обчислень'));
  
  content.push(heading2('32.1. Методологія бенчмаркінгу'));
  content.push(para(
    'Коректний бенчмаркінг паралельних програм — складна задача через JIT warmup, GC паузи, ' +
    'варіабельність планувальника ОС, ефекти кешу. Правила: прогрів JVM (warmup iterations), ' +
    'статистична обробка (середнє + стандартне відхилення), ізоляція від фонових процесів, ' +
    'використання спеціалізованих інструментів (JMH).'
  ));
  
  content.push(heading2('32.2. JMH (Java Microbenchmark Harness)'));
  content.push(code('@BenchmarkMode(Mode.AverageTime)'));
  content.push(code('@OutputTimeUnit(TimeUnit.MILLISECONDS)'));
  content.push(code('@Warmup(iterations = 5, time = 1)'));
  content.push(code('@Measurement(iterations = 10, time = 1)'));
  content.push(code('@Fork(2)'));
  content.push(code('public class MatrixBenchmark {'));
  content.push(code('    @Param({"1", "2", "4", "8"})'));
  content.push(code('    private int threads;'));
  content.push(code('    '));
  content.push(code('    @Benchmark'));
  content.push(code('    public void parallelMultiply(Blackhole bh) {'));
  content.push(code('        bh.consume(multiply(A, B, threads));'));
  content.push(code('    }'));
  content.push(code('}'));
  
  content.push(heading2('32.3. Побудова графіків ефективності'));
  content.push(para(
    'Запускаємо алгоритм при p = 1, 2, 4, 8, 16, 32 потоках, вимірюємо Tₚ. ' +
    'Обчислюємо: Speedup Sₚ = T₁/Tₚ, Efficiency Eₚ = Sₚ/p. ' +
    'Будуємо два графіки: Sₚ від p (порівняти з ідеальною прямою Sₚ=p) і Eₚ від p ' +
    '(ідеал — горизонтальна лінія на 100%). Відхилення вказує на overhead.'
  ));
  
  content.push(heading2('32.4. Аналіз результатів'));
  content.push(bullet('Sublinear speedup (Sₚ < p): норма через overhead синхронізації та комунікацій'));
  content.push(bullet('Plateau або спад при великих p: задача занадто мала або overhead перевищує виграш'));
  content.push(bullet('Superlinear speedup (Sₚ > p): покращена локальність кешу (менше даних на процесор — менше cache miss)'));
  content.push(bullet('"Knee of the curve" — оптимальна кількість потоків для даного розміру задачі'));
  content.push(separator());
  
  // ══════════════════════════════════════════════
  // ПИТАННЯ 34
  // ══════════════════════════════════════════════
  content.push(heading1('Питання 34. Проектування паралельних програм'));
  
  content.push(heading2('34.1. Методологія PCAM (Ian Foster)'));
  content.push(para(
    'Ian Foster запропонував чотирьохетапну методологію проектування паралельних програм PCAM:'
  ));
  content.push(bullet('Partitioning (Розбиття): визначити всі можливі паралельні задачі. Ігнорувати поки що обмеження реальних систем. Мета — максимальний паралелізм.'));
  content.push(bullet('Communication (Комунікація): визначити, які дані потрібно передавати між задачами. Мінімізувати обсяг і частоту комунікацій.'));
  content.push(bullet('Agglomeration (Агломерація): об\'єднати дрібні задачі у більші для зменшення overhead. Врахувати відношення обчислення/комунікація.'));
  content.push(bullet('Mapping (Відображення): призначити агломеровані задачі конкретним процесорам/потокам. Балансувати навантаження і мінімізувати міжпроцесорні комунікації.'));
  
  content.push(heading2('34.2. Принципи ефективного проектування'));
  content.push(bullet('Мінімізувати критичний шлях: виявляти і усувати послідовні залежності'));
  content.push(bullet('Балансування навантаження: рівномірний розподіл роботи (static vs dynamic load balancing)'));
  content.push(bullet('Локальність даних: завдання і дані мають бути на одному вузлі/ядрі'));
  content.push(bullet('Overlap compute and communicate: поки передаються дані, продовжувати обчислення'));
  content.push(bullet('Масштабованість: алгоритм має ефективно працювати при збільшенні p'));
  
  content.push(heading2('34.3. Паттерни паралельного програмування'));
  content.push(bullet('Master-Worker: master роздає задачі, workers виконують і повертають результати'));
  content.push(bullet('Pipeline: послідовні стадії, кожна виконується паралельно на різних даних'));
  content.push(bullet('Producer-Consumer: виробники генерують дані, споживачі обробляють через чергу'));
  content.push(bullet('Map-Reduce: map перетворює кожен елемент незалежно, reduce агрегує результати'));
  content.push(bullet('Divide and Conquer: рекурсивне розбиття задачі (ForkJoin)'));
  content.push(separator());
  
  // ══════════════════════════════════════════════
  // ПИТАННЯ 36
  // ══════════════════════════════════════════════
  content.push(heading1('Питання 36. Моделі пам\'яті паралельних обчислень'));
  
  content.push(heading2('36.1. Послідовна узгодженість (Sequential Consistency)'));
  content.push(para(
    'Найсильніша модель пам\'яті (Lamport, 1979): результат виконання паралельної програми ' +
    'еквівалентний деякому послідовному виконанню тих самих операцій, де операції кожного ' +
    'процесора зберігають свій відносний порядок. Інтуїтивно зрозуміла, але повільна: ' +
    'забороняє більшість оптимізацій компілятора і процесора.'
  ));
  
  content.push(heading2('36.2. TSO (Total Store Order)'));
  content.push(para(
    'Модель пам\'яті x86 процесорів. Записи можуть затримуватись у store buffer. ' +
    'Читання завжди актуальні. Потік бачить власні записи одразу (store forwarding), ' +
    'але інші потоки бачать їх із затримкою. Забезпечує ефективність при низькому рівні ' +
    'невідповідностей для більшості програм.'
  ));
  
  content.push(heading2('36.3. Java Memory Model (JMM)'));
  content.push(para(
    'JMM визначено через відношення happens-before (hb): якщо A hb B, то A видимо перед B. ' +
    'Правила hb: програмний порядок (в одному потоці), unlock→lock того ж монітора, ' +
    'запис volatile→читання того ж поля, start() hb перші дії нового потоку, ' +
    'дії потоку hb join() що повернувся.'
  ));
  content.push(para(
    'JMM дозволяє виконувати all-or-nothing переупорядкування, що не порушують hb. ' +
    'Це дає JIT-компілятору і процесору значну свободу оптимізацій, зберігаючи при цьому ' +
    'коректність при правильному використанні синхронізації.'
  ));
  
  content.push(heading2('36.4. Модель пам\'яті MPI'));
  content.push(para(
    'У MPI кожен процес має власний приватний адресний простір. "Пам\'ять" між процесами ' +
    'не поділяється. Синхронізація стану відбувається лише при обміні повідомленнями: ' +
    'завершення MPI_Send/Recv гарантує, що відправлені дані видимі отримувачу. ' +
    'One-sided MPI (MPI-2/3): RMA (Remote Memory Access) операції Get/Put/Accumulate ' +
    'з явними вікнами синхронізації (MPI_Win_fence, MPI_Win_lock).'
  ));
  content.push(separator());
  
  // ══════════════════════════════════════════════
  // ПИТАННЯ 38
  // ══════════════════════════════════════════════
  content.push(heading1('Питання 38. Стандарт Message Passing Interface (MPI)'));
  
  content.push(heading2('38.1. Що таке MPI'));
  content.push(para(
    'MPI (Message Passing Interface) — стандарт бібліотечного інтерфейсу для програмування ' +
    'паралельних і розподілених систем з передачею повідомлень. Перша версія — MPI-1 (1994), ' +
    'поточна — MPI-4 (2021). MPI не є мовою програмування: це специфікація API, яку реалізують ' +
    'бібліотеки (OpenMPI, MPICH, Intel MPI, Microsoft MPI).'
  ));
  
  content.push(heading2('38.2. Основні концепції'));
  content.push(bullet('Комунікатор (Communicator): контекст комунікації для групи процесів. MPI_COMM_WORLD — всі процеси.'));
  content.push(bullet('Ранг (Rank): унікальний ідентифікатор процесу в комунікаторі (0..size-1).'));
  content.push(bullet('Тег (Tag): числовий ідентифікатор повідомлення для розрізнення типів.'));
  content.push(bullet('Тип даних (Datatype): MPI_INT, MPI_DOUBLE тощо — забезпечує переносимість між архітектурами.'));
  content.push(bullet('Статус (Status): метаінформація отриманого повідомлення (джерело, тег, розмір).'));
  
  content.push(heading2('38.3. Структура MPI програми'));
  content.push(code('MPI.Init(args);'));
  content.push(code('int rank = MPI.COMM_WORLD.Rank();  // ранг поточного процесу'));
  content.push(code('int size = MPI.COMM_WORLD.Size();  // загальна кількість процесів'));
  content.push(code(''));
  content.push(code('if (rank == 0) {'));
  content.push(code('    // root: розподіл задач і збір результатів'));
  content.push(code('} else {'));
  content.push(code('    // workers: виконання підзадач'));
  content.push(code('}'));
  content.push(code(''));
  content.push(code('MPI.Finalize();'));
  
  content.push(heading2('38.4. Запуск MPI програми'));
  content.push(code('# Компіляція та запуск (C/C++):'));
  content.push(code('mpicc -o program program.c'));
  content.push(code('mpirun -np 8 ./program          # 8 процесів на локальній машині'));
  content.push(code('mpirun -np 32 -host node1,node2 ./program  # кластер'));
  content.push(para(
    'Кожен процес виконує той самий бінарний файл, але розгалужується за рангом. ' +
    'JVM MPI реалізується через JNI-обгортки (MPJ Express, FastMPJ) або Java-native MPI binding.'
  ));
  content.push(separator());
  
  // ══════════════════════════════════════════════
  // ПИТАННЯ 40
  // ══════════════════════════════════════════════
  content.push(heading1('Питання 40. Колективні методи обміну повідомленнями MPI'));
  
  content.push(heading2('40.1. Особливості колективних операцій'));
  content.push(para(
    'Колективні операції задіюють всі процеси комунікатора одночасно. ' +
    'Кожен процес повинен викликати колективну операцію (немає поняття "відправник/отримувач" — ' +
    'всі беруть участь). Реалізації оптимізовані під конкретну мережеву топологію ' +
    '(дерево, гіперкуб, butterfly network) і можуть бути значно ефективнішими за ' +
    'ручну реалізацію через point-to-point.'
  ));
  
  content.push(heading2('40.2. Основні колективні операції'));
  content.push(bold('MPI_Bcast — розсилка від root до всіх:'));
  content.push(code('double[] data = (rank == 0) ? loadData() : new double[N];'));
  content.push(code('MPI.COMM_WORLD.Bcast(data, 0, N, MPI.DOUBLE, 0); // root=0'));
  content.push(code('// тепер всі процеси мають data'));
  
  content.push(bold('MPI_Scatter — роздача частин масиву:'));
  content.push(code('double[] allData = (rank == 0) ? fullArray : null;'));
  content.push(code('double[] localData = new double[N/size];'));
  content.push(code('MPI.COMM_WORLD.Scatter(allData, 0, N/size, MPI.DOUBLE,'));
  content.push(code('                      localData, 0, N/size, MPI.DOUBLE, 0);'));
  
  content.push(bold('MPI_Gather — збирання частин до root:'));
  content.push(code('double[] result = (rank == 0) ? new double[N] : null;'));
  content.push(code('MPI.COMM_WORLD.Gather(localData, 0, N/size, MPI.DOUBLE,'));
  content.push(code('                     result, 0, N/size, MPI.DOUBLE, 0);'));
  
  content.push(bold('MPI_Reduce — редукція (сума, max, min тощо):'));
  content.push(code('double[] globalSum = new double[1];'));
  content.push(code('MPI.COMM_WORLD.Reduce(new double[]{localSum}, 0, globalSum, 0,'));
  content.push(code('                     1, MPI.DOUBLE, MPI.SUM, 0);'));
  
  content.push(bold('MPI_Allreduce — редукція з розсилкою результату ВСІМ:'));
  content.push(code('double[] globalMax = new double[1];'));
  content.push(code('MPI.COMM_WORLD.Allreduce(new double[]{localMax}, 0, globalMax, 0,'));
  content.push(code('                        1, MPI.DOUBLE, MPI.MAX);'));
  
  content.push(bold('MPI_Allgather — кожен отримує повний зібраний масив:'));
  content.push(code('MPI.COMM_WORLD.Allgather(localPart, 0, chunk, MPI.DOUBLE,'));
  content.push(code('                        fullArray, 0, chunk, MPI.DOUBLE);'));
  
  content.push(bold('MPI_Barrier — синхронізаційний бар\'єр:'));
  content.push(code('MPI.COMM_WORLD.Barrier(); // всі чекають тут'));
  
  content.push(heading2('40.3. Застосування для множення матриць'));
  content.push(para(
    'Алгоритм стрічкового множення з колективними операціями: root розподіляє рядки A через ' +
    'Scatter, всі процеси отримують повну B через Bcast, виконують локальне множення, ' +
    'root збирає рядки C через Gather. Складність комунікацій: O(n²/p) — ефективно.'
  ));
  content.push(separator());
  
  // ══════════════════════════════════════════════
  // ПИТАННЯ 42
  // ══════════════════════════════════════════════
  content.push(heading1('Питання 42. Розробка ефективних паралельних алгоритмів в OpenMPI'));
  
  content.push(heading2('42.1. Принципи ефективності в MPI'));
  content.push(bullet('Мінімізувати кількість і розмір повідомлень: об\'єднувати малі повідомлення у великі (message coalescing)'));
  content.push(bullet('Overlap computation and communication: використовувати неблокуючі Isend/Irecv, поки передаються дані — рахувати'));
  content.push(bullet('Колективні операції замість ручних циклів: реалізації MPI оптимізовані під топологію'));
  content.push(bullet('Топологічна оптимізація: MPI_Cart_create для 2D/3D сіток, MPI_Graph для довільних топологій'));
  content.push(bullet('Derived datatypes: MPI_Type_vector для не-суміжних даних (стовпці матриці) без копіювання'));
  
  content.push(heading2('42.2. Неблокуючі комунікації для overlap'));
  content.push(code('// Без overlap:'));
  content.push(code('compute(localData);              // обчислення'));
  content.push(code('MPI.Send(borderData, ...);       // потім передача'));
  content.push(code('MPI.Recv(neighborData, ...);'));
  content.push(code(''));
  content.push(code('// З overlap (ефективніше):'));
  content.push(code('Request sendReq = MPI.COMM_WORLD.Isend(borderData, ...); // запустити передачу'));
  content.push(code('Request recvReq = MPI.COMM_WORLD.Irecv(neighborData, ...);'));
  content.push(code('compute(innerData);  // рахувати внутрішні дані, поки передаються межі'));
  content.push(code('Request.Waitall(new Request[]{sendReq, recvReq}); // дочекатись'));
  content.push(code('update(borderData, neighborData); // оновити межові елементи'));
  
  content.push(heading2('42.3. Derived Datatypes'));
  content.push(code('// Відправити стовпець матриці n×n без копіювання:'));
  content.push(code('MPI.Datatype columnType = MPI.DOUBLE.Vector(n, 1, n);'));
  content.push(code('// n блоків по 1 елементу з кроком n (рядок)'));
  content.push(code('columnType.Commit();'));
  content.push(code('MPI.COMM_WORLD.Send(matrix, col, 1, columnType, dest, tag);'));
  content.push(code('columnType.Free();'));
  
  content.push(heading2('42.4. Профілювання MPI програм'));
  content.push(para(
    'Vampir, TAU, MPE — інструменти трасування MPI. Будують діаграми Ганта: кожна рядок — ' +
    'процес, смуги — обчислення, лінії між рядками — повідомлення. Дозволяють виявити: ' +
    'load imbalance (різна довжина смуг), communication bottleneck (багато ліній), ' +
    'barrier inefficiency (довге очікування на бар\'єрі).'
  ));
  content.push(separator());
  
  // ══════════════════════════════════════════════
  // ПИТАННЯ 44
  // ══════════════════════════════════════════════
  content.push(heading1('Питання 44. Паралельна реалізація алгоритму МГУА'));
  
  content.push(heading2('44.1. Що таке МГУА'));
  content.push(para(
    'МГУА (Метод Групового Урахування Аргументів, Group Method of Data Handling — GMDH) — ' +
    'алгоритм машинного навчання, запропонований О.Г. Івахненком (1968, Київ). ' +
    'Автоматично будує поліноміальні регресійні моделі через послідовний відбір найкращих ' +
    'комбінацій ознак. Широко застосовується у прогнозуванні, ідентифікації систем, ' +
    'аналізі даних.'
  ));
  
  content.push(heading2('44.2. Структура алгоритму'));
  content.push(bullet('Шар 0: вхідні змінні x₁, x₂, ..., xₙ'));
  content.push(bullet('Шар k: для кожної пари (xᵢ, xⱼ) будується поліном Колмогорова-Габора: y = a + b·xᵢ + c·xⱼ + d·xᵢ² + e·xⱼ² + f·xᵢ·xⱼ'));
  content.push(bullet('Відбір: залишаються M найкращих часткових описів за критерієм (наприклад, MSE на тестовій вибірці)'));
  content.push(bullet('Повторювати шари до зупинення (критерій регулярності перестає покращуватись)'));
  
  content.push(heading2('44.3. Потенціал паралелізму'));
  content.push(para(
    'На кожному шарі C(n, 2) = n*(n-1)/2 пар незалежно будуються і оцінюються. ' +
    'Всі ці задачі повністю незалежні — ідеальна Data Parallelism ситуація. ' +
    'Прискорення лінійне: Sₚ ≈ p аж до кількості пар.'
  ));
  
  content.push(heading2('44.4. Реалізація в Java'));
  content.push(code('ExecutorService pool = Executors.newFixedThreadPool(nCores);'));
  content.push(code('List<Future<Model>> futures = new ArrayList<>();'));
  content.push(code(''));
  content.push(code('// Для кожного шару: паралельно будуємо всі пари'));
  content.push(code('for (int i = 0; i < inputs.size(); i++) {'));
  content.push(code('    for (int j = i + 1; j < inputs.size(); j++) {'));
  content.push(code('        final int fi = i, fj = j;'));
  content.push(code('        futures.add(pool.submit(() ->'));
  content.push(code('            fitPolynomial(inputs.get(fi), inputs.get(fj), trainData)));'));
  content.push(code('    }'));
  content.push(code('}'));
  content.push(code(''));
  content.push(code('// Зібрати результати і відібрати найкращі M:'));
  content.push(code('List<Model> models = futures.stream()'));
  content.push(code('    .map(f -> { try { return f.get(); } catch (Exception e) { ... } })'));
  content.push(code('    .sorted(Comparator.comparingDouble(Model::getMse))'));
  content.push(code('    .limit(M)'));
  content.push(code('    .collect(Collectors.toList());'));
  content.push(separator());
  
  // ══════════════════════════════════════════════
  // ПИТАННЯ 46
  // ══════════════════════════════════════════════
  content.push(heading1('Питання 46. Технологія Remote Method Invocation (RMI)'));
  
  content.push(heading2('46.1. Що таке Java RMI'));
  content.push(para(
    'RMI (Remote Method Invocation) — технологія Java для виклику методів об\'єктів, ' +
    'що знаходяться в іншій JVM (на іншій машині або тому ж хості). ' +
    'Є Java-специфічною реалізацією концепції RPC (Remote Procedure Call). ' +
    'Прозорість: виклик віддаленого методу виглядає (майже) як виклик локального.'
  ));
  
  content.push(heading2('46.2. Архітектура RMI'));
  content.push(bullet('Remote Interface: Java-інтерфейс, що розширює java.rmi.Remote, оголошує методи що кидають RemoteException'));
  content.push(bullet('Remote Object (Servant): реалізація інтерфейсу на сервері, розширює UnicastRemoteObject'));
  content.push(bullet('Stub: клієнтський проксі, серіалізує аргументи і відправляє через мережу'));
  content.push(bullet('Skeleton (до Java 2): серверний компонент, десеріалізує і викликає реальний об\'єкт'));
  content.push(bullet('RMI Registry: іменний сервіс (bind/lookup) для пошуку віддалених об\'єктів за іменем'));
  
  content.push(heading2('46.3. Приклад реалізації'));
  content.push(code('// 1. Remote Interface:'));
  content.push(code('public interface MatrixService extends Remote {'));
  content.push(code('    double[][] multiply(double[][] A, double[][] B)'));
  content.push(code('        throws RemoteException;'));
  content.push(code('}'));
  content.push(code(''));
  content.push(code('// 2. Server implementation:'));
  content.push(code('public class MatrixServiceImpl extends UnicastRemoteObject'));
  content.push(code('        implements MatrixService {'));
  content.push(code('    public double[][] multiply(double[][] A, double[][] B) {'));
  content.push(code('        return parallelMultiply(A, B);'));
  content.push(code('    }'));
  content.push(code('}'));
  content.push(code(''));
  content.push(code('// 3. Server main:'));
  content.push(code('MatrixService service = new MatrixServiceImpl();'));
  content.push(code('LocateRegistry.createRegistry(1099);'));
  content.push(code('Naming.rebind("rmi://localhost/MatrixService", service);'));
  content.push(code(''));
  content.push(code('// 4. Client:'));
  content.push(code('MatrixService svc = (MatrixService)'));
  content.push(code('    Naming.lookup("rmi://server-host/MatrixService");'));
  content.push(code('double[][] C = svc.multiply(A, B); // прозорий віддалений виклик'));
  
  content.push(heading2('46.4. RMI vs сучасні альтернативи'));
  content.push(bullet('RMI: Java-only, проста в налаштуванні, погана продуктивність для великих даних'));
  content.push(bullet('gRPC: поліглотний, Protocol Buffers, HTTP/2, стримінг, висока продуктивність'));
  content.push(bullet('REST/HTTP: найпростіший, повсюдна підтримка, але більший overhead'));
  content.push(bullet('Apache Thrift: Facebook, поліглотний, ефективний бінарний протокол'));
  content.push(separator());
  
  // ══════════════════════════════════════════════
  // ПИТАННЯ 48
  // ══════════════════════════════════════════════
  content.push(heading1('Питання 48. Базові складові грід-системи'));
  
  content.push(heading2('48.1. Компоненти грід-інфраструктури'));
  content.push(para(
    'Грід-система складається з кількох ключових компонентів, кожен з яких виконує ' +
    'специфічну роль у забезпеченні прозорого доступу до розподілених ресурсів:'
  ));
  
  content.push(heading2('48.2. Ресурсні вузли (Resource Nodes)'));
  content.push(para(
    'Фізичні або віртуальні обчислювальні ресурси: суперкомп\'ютери, кластери, ' +
    'робочі станції, системи зберігання даних. Кожен вузол має локальну ОС і ' +
    'локальний менеджер ресурсів (PBS, SLURM, LSF). Grid middleware взаємодіє ' +
    'з локальними менеджерами через стандартизовані інтерфейси (GRAM, UNICORE).'
  ));
  
  content.push(heading2('48.3. Безпека (Security Infrastructure)'));
  content.push(para(
    'GSI (Grid Security Infrastructure): основа безпеки Globus. Використовує X.509 ' +
    'сертифікати і PKI для автентифікації. Proxy certificates: короткочасні сертифікати ' +
    'для делегування прав без постійного введення паролю. Single Sign-On через MyProxy: ' +
    'зберігання проксі сертифікатів на сервері.'
  ));
  
  content.push(heading2('48.4. Управління ресурсами (Resource Management)'));
  content.push(bullet('GRAM (Globus Resource Allocation and Management): подання задач на виконання, моніторинг стану'));
  content.push(bullet('MDS (Monitoring and Discovery Service): реєстр доступних ресурсів і їх характеристик'));
  content.push(bullet('GRIS (Grid Resource Information Service): публікація інформації про локальні ресурси'));
  content.push(bullet('GIIS (Grid Index Information Service): агрегатор інформації з кількох GRIS'));
  
  content.push(heading2('48.5. Передача даних (Data Management)'));
  content.push(para(
    'GridFTP: оптимізований протокол передачі файлів для Grid. Підтримує: паралельні потоки ' +
    '(збільшення пропускної здатності), часткові передачі, відновлення після збою, ' +
    'узгоджену безпеку з GSI. Replica Management Service: управління копіями даних ' +
    '(реплікація на кілька вузлів для відмовостійкості і локальності доступу).'
  ));
  content.push(separator());
  
  // ══════════════════════════════════════════════
  // ПИТАННЯ 50
  // ══════════════════════════════════════════════
  content.push(heading1('Питання 50. Організація і управління розподіленими грід-ресурсами'));
  
  content.push(heading2('50.1. Планування задач у Grid'));
  content.push(para(
    'Grid scheduling — процес призначення задач на ресурси Grid. На відміну від локального ' +
    'планування, Grid scheduler має враховувати: гетерогенність ресурсів (різна продуктивність), ' +
    'мережеву топологію (перенесення даних), завантаженість (загрузка вузлів), ' +
    'вартість (платні ресурси), пріоритети Virtual Organization.'
  ));
  
  content.push(heading2('50.2. Рівні планування'));
  content.push(bullet('Local Resource Manager (LRM): планує задачі на одному вузлі/кластері (PBS Torque, SLURM, LSF). Знає реальний стан локальних ресурсів.'));
  content.push(bullet('Meta-Scheduler / Grid Broker: вибирає, на якому ресурсному вузлі виконувати задачу. Алгоритми: min-min, max-min, sufferage. Приклади: Condor-G, GridWay.'));
  content.push(bullet('Workflow Engine: управляє складними залежностями між задачами (DAG workflow). Приклади: Pegasus, Taverna.'));
  
  content.push(heading2('50.3. SLA і QoS'));
  content.push(para(
    'SLA (Service Level Agreement) в Grid: угода між користувачем і провайдером ресурсів про ' +
    'гарантовані характеристики (термін виконання, доступна пам\'ять, пропускна здатність). ' +
    'WSRF (Web Services Resource Framework) — стандарт для опису і управління статефулими ' +
    'Web сервісами в Grid. OGSA (Open Grid Services Architecture) — архітектурний фреймворк ' +
    'на основі Web Services для Grid.'
  ));
  
  content.push(heading2('50.4. Відмовостійкість'));
  content.push(bullet('Checkpointing: збереження стану задачі для відновлення після відмови вузла'));
  content.push(bullet('Replication: повторний запуск задачі на іншому вузлі при відмові'));
  content.push(bullet('Work unit validation: BOINC відправляє одну задачу кільком вузлам і порівнює результати'));
  content.push(bullet('Heartbeat monitoring: регулярні сигнали від вузлів; відсутність сигналу → вузол вважається недоступним'));
  
  content.push(heading2('50.5. Сучасний стан і Cloud vs Grid'));
  content.push(para(
    'Хмарні технології (AWS, Azure, GCP) значною мірою витіснили Grid у комерційному секторі: ' +
    'простіша моделі оплати, кращий інтерфейс, автоматичне масштабування. ' +
    'Grid залишається важливим у науковому середовищі (CERN LHC Grid, EGI, XSEDE у США), ' +
    'де потрібно об\'єднати ресурси незалежних організацій без єдиного власника. ' +
    'Тенденція: гібридні Grid+Cloud системи, де Grid оркеструє ресурси хмарних провайдерів.'
  ));
  content.push(separator());
  
  // ─────────────────────────────────────────────
  // ПОБУДОВА ДОКУМЕНТУ
  // ─────────────────────────────────────────────
  
  const doc = new Document({
    styles: {
      default: {
        document: { run: { font: 'Arial', size: 24 } },
      },
      paragraphStyles: [
        {
          id: 'Heading1',
          name: 'Heading 1',
          basedOn: 'Normal',
          next: 'Normal',
          quickFormat: true,
          run: { size: 32, bold: true, font: 'Arial', color: '2E4057' },
          paragraph: { spacing: { before: 480, after: 200 }, outlineLevel: 0 },
        },
        {
          id: 'Heading2',
          name: 'Heading 2',
          basedOn: 'Normal',
          next: 'Normal',
          quickFormat: true,
          run: { size: 26, bold: true, font: 'Arial', color: '2E75B6' },
          paragraph: { spacing: { before: 300, after: 120 }, outlineLevel: 1 },
        },
      ],
    },
    numbering: {
      config: [
        {
          reference: 'bullets',
          levels: [
            {
              level: 0,
              format: LevelFormat.BULLET,
              text: '\u2022',
              alignment: AlignmentType.LEFT,
              style: { paragraph: { indent: { left: 720, hanging: 360 } } },
            },
          ],
        },
      ],
    },
    sections: [
      {
        properties: {
          page: {
            size: { width: 11906, height: 16838 }, // A4
            margin: { top: 1134, right: 850, bottom: 1134, left: 1134 },
          },
        },
        children: [
          new Paragraph({
            children: [new TextRun({ text: '', size: 24 })],
            spacing: { before: 2000, after: 400 },
          }),
          new Paragraph({
            children: [new TextRun({ text: 'Технологія паралельних обчислень', bold: true, size: 48, font: 'Arial', color: '2E4057' })],
            alignment: AlignmentType.CENTER,
            spacing: { before: 0, after: 400 },
          }),
          new Paragraph({
            children: [new TextRun({ text: 'Відповіді на екзаменаційні питання — Частина 2', size: 32, font: 'Arial', color: '666666' })],
            alignment: AlignmentType.CENTER,
            spacing: { before: 0, after: 200 },
          }),
          new Paragraph({
            children: [new TextRun({ text: 'Мова програмування: Java  |  Питання: 2, 4, 6, 8, 10, 12, 14, 16, 18, 20, 22, 24, 26, 28, 30, 32, 34, 36, 38, 40, 42, 44, 46, 48, 50', size: 22, font: 'Arial', color: '888888' })],
            alignment: AlignmentType.CENTER,
            spacing: { before: 0, after: 3000 },
          }),
          new Paragraph({ children: [new PageBreak()] }),
          ...content,
        ],
      },
    ],
  });
  
  Packer.toBuffer(doc).then((buffer) => {
    fs.writeFileSync('./docs/ev3.docx', buffer);
  });