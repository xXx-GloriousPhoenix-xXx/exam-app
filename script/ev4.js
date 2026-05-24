// ╔══════════════════════════════════════════════════════════════════╗
// ║  Технологія паралельних обчислень — парні питання (2..50)       ║
// ║  npm install docx   →   node generate_exam_even.js              ║
// ╚══════════════════════════════════════════════════════════════════╝

const {
    Document, Packer, Paragraph, TextRun, HeadingLevel,
    AlignmentType, LevelFormat, PageNumber,
    Header, Footer, BorderStyle, WidthType,
    Table, TableRow, TableCell, ShadingType,
  } = require('docx');
  const fs = require('fs');
  
  // ─── helpers ────────────────────────────────────────────────────────────────
  
  const A4_W = 11906, A4_H = 16838;
  const MARGIN = 1440;
  const CONTENT_W = A4_W - MARGIN * 2; // 9026 DXA
  
  function h1(text) {
    return new Paragraph({
      heading: HeadingLevel.HEADING_1,
      pageBreakBefore: true,
      children: [new TextRun({ text, bold: true, size: 32, font: 'Arial', color: '1F3864' })],
      spacing: { before: 320, after: 160 },
    });
  }
  
  function h2(text) {
    return new Paragraph({
      heading: HeadingLevel.HEADING_2,
      children: [new TextRun({ text, bold: true, size: 28, font: 'Arial', color: '2E75B6' })],
      spacing: { before: 220, after: 100 },
    });
  }
  
  function p(text) {
    return new Paragraph({
      alignment: AlignmentType.JUSTIFIED,
      children: [new TextRun({ text, size: 24, font: 'Arial' })],
      spacing: { before: 80, after: 80 },
    });
  }
  
  function bullet(text) {
    return new Paragraph({
      numbering: { reference: 'bullets', level: 0 },
      children: [new TextRun({ text, size: 24, font: 'Arial' })],
      spacing: { before: 40, after: 40 },
    });
  }
  
  function code(line) {
    return new Paragraph({
      children: [new TextRun({ text: line, font: 'Courier New', size: 20, color: '1A3A5C' })],
      spacing: { before: 40, after: 40 },
      indent: { left: 720 },
      border: { left: { style: BorderStyle.SINGLE, size: 6, color: '2E75B6', space: 8 } },
    });
  }
  
  function codes(lines) { return lines.map(code); }
  
  // ─── content ────────────────────────────────────────────────────────────────
  
  const body = [];
  
  // ═══════════════════════════════════════════════════════════════════
  // ПИТАННЯ 2
  // ═══════════════════════════════════════════════════════════════════
  body.push(h1('Питання 2. Класифікація паралельних обчислювальних систем'));
  
  body.push(h2('2.1. Таксономія Фліна'));
  body.push(p('Найпоширенішою класифікацією є таксономія Фліна (Flynn, 1966), що ділить архітектури за кількістю потоків інструкцій (I) та потоків даних (D):'));
  body.push(bullet('SISD (Single Instruction / Single Data) — послідовна архітектура фон Неймана. Один процесор виконує одну інструкцію над одним операндом. Приклад: класичний Intel 8086.'));
  body.push(bullet('SIMD (Single Instruction / Multiple Data) — одна інструкція застосовується одночасно до вектора даних. Типово: GPU (тисячі ядер з однаковим шейдером), CPU-розширення SSE/AVX. В Java: parallelStream() на масиві, де кожен елемент обробляється однаково.'));
  body.push(bullet('MISD (Multiple Instruction / Single Data) — кілька процесорів виконують різні операції над тими самими даними. Рідкісна: відмовостійкі системи (Space Shuttle), де кілька незалежних комп\'ютерів перевіряють один і той самий сенсорний потік.'));
  body.push(bullet('MIMD (Multiple Instruction / Multiple Data) — кожен процесор виконує власну програму над своїми даними. Сучасні багатоядерні CPU, кластери, суперкомп\'ютери. Поділяється на підкатегорії за організацією пам\'яті.'));
  
  body.push(h2('2.2. Класифікація за організацією пам\'яті'));
  body.push(bullet('Shared Memory (спільна пам\'ять): всі процесори мають доступ до єдиного адресного простору через шину або мережу з\'єднань. Спрощує програмування (глобальні змінні, Java-потоки), але обмежує масштабованість через конкуренцію за пам\'ять та шину.'));
  body.push(bullet('Distributed Memory (розподілена пам\'ять): кожен вузол має приватну пам\'ять. Взаємодія — виключно через передачу повідомлень (MPI). Масштабується до тисяч вузлів. Програмувати складніше.'));
  body.push(bullet('Distributed Shared Memory (DSM): фізично розподілена, логічно — спільна. ОС або апаратура підтримують ілюзію єдиного адресного простору. Приклад: NUMA-системи (Non-Uniform Memory Access).'));
  
  body.push(h2('2.3. Класифікація за зв\'язністю'));
  body.push(bullet('Тісно зв\'язані (Tightly Coupled): процесори на спільній шині або кросбарі. Низька затримка, обмежена кількість вузлів (типово до 64).'));
  body.push(bullet('Слабо зв\'язані (Loosely Coupled): вузли з\'єднані мережею (InfiniBand, Ethernet). Масштабуються до тисяч вузлів. Вища затримка комунікацій.'));
  
  body.push(h2('2.4. Класифікація MPP, SMP, Cluster'));
  body.push(bullet('SMP (Symmetric Multiprocessing): 2–64 процесори зі спільною пам\'яттю та єдиною ОС. Типовий сервер. Прозоро для Java-потоків.'));
  body.push(bullet('MPP (Massively Parallel Processing): тисячі вузлів, кожен — SMP. Власна ОС на кожному. Зв\'язок через MPI.'));
  body.push(bullet('Cluster: стандартні сервери, з\'єднані мережею. Commodity hardware. Hadoop, Spark, OpenMPI.'));
  body.push(bullet('NUMA: неоднорідна затримка доступу до пам\'яті. Доступ до локальної пам\'яті швидший, ніж до пам\'яті іншого сокету. Важливо враховувати при плануванні потоків.'));
  
  
  // ═══════════════════════════════════════════════════════════════════
  // ПИТАННЯ 4
  // ═══════════════════════════════════════════════════════════════════
  body.push(h1('Питання 4. Системи із загальною та розподіленою пам\'яттю'));
  
  body.push(h2('4.1. Системи із загальною пам\'яттю'));
  body.push(p('У системах із загальною пам\'яттю (Shared Memory Systems) всі процесори мають доступ до єдиного фізичного адресного простору. Комунікація між процесорами відбувається через читання та запис спільних змінних. Синхронізація — через mutex, семафори, бар\'єри.'));
  body.push(p('Переваги: простота програмування (Java Thread, OpenMP), відсутність явної передачі даних, низька затримка комунікації (пам\'ять). Недоліки: обмежена масштабованість (шина пам\'яті стає вузьким місцем), складна синхронізація, cache coherence overhead.'));
  body.push(p('Cache Coherence: при записі одним процесором кеш-ліній, інші процесори мають побачити оновлене значення. Протокол MESI (Modified, Exclusive, Shared, Invalid) забезпечує узгодженість. Java Memory Model (JMM) гарантує видимість через volatile, synchronized, happens-before.'));
  body.push(...codes([
    '// Java — спільна пам\'ять між потоками',
    'class SharedCounter {',
    '    private volatile int count = 0; // volatile — видимість між потоками',
    '    public synchronized void increment() { count++; }',
    '    public int get() { return count; }',
    '}',
  ]));
  
  body.push(h2('4.2. NUMA (Non-Uniform Memory Access)'));
  body.push(p('Сучасні багатосокетні сервери використовують NUMA: пам\'ять фізично розподілена між вузлами (сокетами), але логічно — єдиний адресний простір. Доступ до локальної пам\'яті: 4–8 нс, до віддаленої: 20–40 нс. Java-програміст не бачить цього явно, але CPU-affinity та numactl дозволяють оптимізувати.'));
  
  body.push(h2('4.3. Системи з розподіленою пам\'яттю'));
  body.push(p('У системах із розподіленою пам\'яттю (Distributed Memory Systems) кожен вузол має власну приватну пам\'ять. Немає спільних змінних — вся комунікація явна, через передачу повідомлень. Стандарт: MPI (Message Passing Interface).'));
  body.push(p('Переваги: лінійне масштабування (немає боротьби за шину), масштаб до сотень тисяч ядер (Frontier: 9 млн ядер). Недоліки: складність програмування, необхідність ручного розподілу даних, затримки мережі (InfiniBand: ~1 мкс, Ethernet: ~10–100 мкс).'));
  body.push(...codes([
    '// MPI-псевдокод для розподіленої пам\'яті',
    'MPI_Init(&argc, &argv);',
    'MPI_Comm_rank(MPI_COMM_WORLD, &rank); // хто я?',
    'MPI_Comm_size(MPI_COMM_WORLD, &size); // скільки нас?',
    'if (rank == 0) MPI_Send(data, n, MPI_DOUBLE, 1, 0, MPI_COMM_WORLD);',
    'else           MPI_Recv(buf,  n, MPI_DOUBLE, 0, 0, MPI_COMM_WORLD, &status);',
    'MPI_Finalize();',
  ]));
  
  body.push(h2('4.4. Порівняльна характеристика'));
  body.push(p('Спільна пам\'ять: зручна для задач зі складними структурами даних, підходить до ~64 ядер. Розподілена: масштабується до мільйонів ядер, підходить для регулярних задач (матриці, стенцильні обчислення). Гібридна (MPI + OpenMP або MPI + Java Thread): кожен вузол — спільна пам\'ять всередині, між вузлами — MPI.'));
  
  
  // ═══════════════════════════════════════════════════════════════════
  // ПИТАННЯ 6
  // ═══════════════════════════════════════════════════════════════════
  body.push(h1('Питання 6. Багатопоточна технологія Java'));
  
  body.push(h2('6.1. Потоки в Java — основи'));
  body.push(p('Java з версії 1.0 підтримує багатопоточність на рівні мови та JVM. Кожна програма стартує з головного потоку (main thread). Новий потік створюється двома основними способами:'));
  body.push(bullet('Успадкування від java.lang.Thread та перевизначення run().'));
  body.push(bullet('Реалізація інтерфейсу java.lang.Runnable (рекомендований підхід — уникає обмежень одиночного успадкування).'));
  body.push(p('Починаючи з Java 5, з\'явився пакет java.util.concurrent (JUC) — повноцінна бібліотека для паралельного програмування. Java 8 додала parallelStream() та CompletableFuture. Java 19–21 — Virtual Threads (Project Loom).'));
  
  body.push(h2('6.2. Способи створення потоків'));
  body.push(...codes([
    '// Спосіб 1: Thread',
    'class Compute extends Thread {',
    '    public void run() { System.out.println("Thread: " + getName()); }',
    '}',
    'new Compute().start();',
    '',
    '// Спосіб 2: Runnable (лямбда)',
    'Thread t = new Thread(() -> System.out.println("Runnable thread"));',
    't.start();',
    '',
    '// Спосіб 3: Callable + Future',
    'ExecutorService ex = Executors.newFixedThreadPool(4);',
    'Future<Double> f = ex.submit(() -> Math.sqrt(2.0));',
    'System.out.println(f.get()); // 1.4142...',
    'ex.shutdown();',
  ]));
  
  body.push(h2('6.3. Java Memory Model (JMM)'));
  body.push(p('JMM визначає, як потоки взаємодіють через пам\'ять. Ключова концепція: happens-before. Якщо дія A happens-before дії B, то результати A видимі B. Гарантії happens-before:'));
  body.push(bullet('Запис у volatile-поле happens-before читання того самого поля.'));
  body.push(bullet('Виклик start() happens-before будь-якої дії у запущеному потоці.'));
  body.push(bullet('Всі дії потоку happens-before join() на цьому потоці.'));
  body.push(bullet('Вихід з synchronized-блоку happens-before входу в той самий synchronized-блок іншим потоком.'));
  
  body.push(h2('6.4. Паралельні стріми (parallelStream)'));
  body.push(...codes([
    'List<Integer> nums = IntStream.rangeClosed(1, 10_000_000)',
    '                              .boxed().collect(Collectors.toList());',
    '',
    '// Паралельне сумування — автоматично ділить на потоки',
    'long sum = nums.parallelStream()',
    '               .mapToLong(Integer::longValue)',
    '               .sum();',
    '// Використовує ForkJoinPool.commonPool() внутрішньо',
  ]));
  
  body.push(h2('6.5. Virtual Threads (Java 21+)'));
  body.push(p('Project Loom: мільйони легковагих virtual threads, що маплюються на невелику кількість platform threads. Ідеально для I/O-bound задач: кожен HTTP-запит у своєму virtual thread без overhead. Синтаксис: Thread.ofVirtual().start(runnable).'));
  
  
  // ═══════════════════════════════════════════════════════════════════
  // ПИТАННЯ 8
  // ═══════════════════════════════════════════════════════════════════
  body.push(h1('Питання 8. Стани потоку та переходи між станами'));
  
  body.push(h2('8.1. Шість станів потоку Java'));
  body.push(p('Перерахування Thread.State визначає шість станів:'));
  body.push(bullet('NEW — потік створений (new Thread()), але start() ще не викликано.'));
  body.push(bullet('RUNNABLE — потік виконується або готовий до виконання (чекає на CPU від планувальника ОС). Включає обидва стани ОС: running та ready.'));
  body.push(bullet('BLOCKED — потік чекає на монітор об\'єкта, щоб увійти в synchronized-метод або блок, який вже захопив інший потік.'));
  body.push(bullet('WAITING — потік нескінченно чекає сигналу від іншого потоку. Причини: Object.wait(), Thread.join() без таймауту, LockSupport.park().'));
  body.push(bullet('TIMED_WAITING — потік чекає обмежений час. Причини: Thread.sleep(millis), Object.wait(millis), Thread.join(millis), LockSupport.parkNanos().'));
  body.push(bullet('TERMINATED — метод run() завершився (нормально або через виключення). Потік не можна перезапустити.'));
  
  body.push(h2('8.2. Переходи між станами'));
  body.push(p('NEW -> RUNNABLE: виклик start().'));
  body.push(p('RUNNABLE -> BLOCKED: потік намагається увійти в synchronized, але монітор зайнятий.'));
  body.push(p('BLOCKED -> RUNNABLE: монітор звільнився, потік отримав його.'));
  body.push(p('RUNNABLE -> WAITING: виклик wait(), join(), park().'));
  body.push(p('WAITING -> RUNNABLE: виклик notify()/notifyAll() або join()-потік завершився.'));
  body.push(p('RUNNABLE -> TIMED_WAITING: sleep(n), wait(n), join(n).'));
  body.push(p('TIMED_WAITING -> RUNNABLE: час вийшов або отримали notify().'));
  body.push(p('RUNNABLE -> TERMINATED: run() завершився або кинуто непійманий виняток.'));
  
  body.push(h2('8.3. Перевірка стану'));
  body.push(...codes([
    'Thread t = new Thread(() -> {',
    '    try { Thread.sleep(500); } catch (InterruptedException e) {}',
    '});',
    'System.out.println(t.getState()); // NEW',
    't.start();',
    'Thread.sleep(100);',
    'System.out.println(t.getState()); // TIMED_WAITING',
    't.join();',
    'System.out.println(t.getState()); // TERMINATED',
  ]));
  
  body.push(h2('8.4. Важливі особливості'));
  body.push(p('BLOCKED та WAITING — принципово різні стани: BLOCKED чекає монітора (конкуренція за synchronized), WAITING чекає явного сповіщення. Розрізнення важливе для діагностики проблем: deadlock зазвичай виявляється у стані BLOCKED, а starvation — у тривалому WAITING.'));
  
  
  // ═══════════════════════════════════════════════════════════════════
  // ПИТАННЯ 10
  // ═══════════════════════════════════════════════════════════════════
  body.push(h1('Питання 10. Алгоритми паралельного множення матриць'));
  
  body.push(h2('10.1. Послідовне множення та складність'));
  body.push(p('C = A × B, де A(n×m), B(m×k), C(n×k). Елемент C[i][j] = Σ A[i][t]×B[t][j] (t = 0..m-1). Послідовна складність O(n³) для квадратних матриць. Ціль паралелізації: рівномірно розподілити O(n³) операцій між p процесорами.'));
  
  body.push(h2('10.2. Стрічковий (Row-Striped) алгоритм'));
  body.push(p('Матриця A ділиться на горизонтальні стрічки (рядки). Кожен потік/процес отримує свою стрічку і обчислює відповідні рядки C. Матриця B — спільна (read-only). Найпростіший підхід, добре підходить для спільної пам\'яті (Java Thread).'));
  body.push(...codes([
    'int chunk = n / numThreads;',
    'for (int t = 0; t < numThreads; t++) {',
    '    final int from = t * chunk;',
    '    final int to = (t == numThreads-1) ? n : from + chunk;',
    '    pool.submit(() -> {',
    '        for (int i = from; i < to; i++)',
    '            for (int j = 0; j < k; j++) {',
    '                C[i][j] = 0;',
    '                for (int r = 0; r < m; r++) C[i][j] += A[i][r] * B[r][j];',
    '            }',
    '    });',
    '}',
  ]));
  
  body.push(h2('10.3. Блочний алгоритм (Block Decomposition)'));
  body.push(p('Матриці A, B, C розбиваються на блоки розміром (n/q)×(n/q), де q×q = p (кількість процесів). Кожен процес (i,j) обчислює блок C[i][j] = Σ A[i][k]×B[k][j]. Краща кеш-поведінка, ніж рядковий. Основа для алгоритмів Фокса та Кеннона.'));
  
  body.push(h2('10.4. Алгоритм Кеннона (Cannon)'));
  body.push(p('Розроблений для тороїдальної сітки q×q процесів. Початковий зсув рядків A та стовпців B. Потім q кроків: локальне множення блоків та кільцевий зсув (A — вліво, B — вгору). Мінімальна комунікація: O(n²/q) на крок. Складність: O(n³/p) обч. + O(n²/√p) комунік.'));
  
  body.push(h2('10.5. Алгоритм Фокса'));
  body.push(p('Також для сітки q×q. На кожному з q кроків k: процес (i,k) надсилає broadcast блоку A[i][k] всім у своєму рядку. Кожен процес (i,j) множить отриманий блок A на свій B[k][j]. Потім B-блоки зсуваються вгору. Використовує колективні операції MPI_Bcast.'));
  
  body.push(h2('10.6. Аналіз прискорення'));
  body.push(p('Теоретичне прискорення S = p. На практиці: S(4) ≈ 3.6–3.9, S(16) ≈ 13–15. Основні втрати: синхронізація між кроками, передача даних по мережі, незбалансованість при непропорційному n/p.'));
  
  
  // ═══════════════════════════════════════════════════════════════════
  // ПИТАННЯ 12
  // ═══════════════════════════════════════════════════════════════════
  body.push(h1('Питання 12. Алгоритм Фокса паралельного множення матриць'));
  
  body.push(h2('12.1. Постановка та ідея'));
  body.push(p('Алгоритм Фокса (Fox\'s Algorithm, 1987) призначений для квадратних матриць n×n на сітці q×q процесів (p = q²). Кожен процес (i,j) спочатку зберігає блоки A_ij та B_ij розміром (n/q)×(n/q). Результат: кожен процес накопичує блок C_ij = Σ A_ik × B_kj (k=0..q-1).'));
  
  body.push(h2('12.2. Покрокове виконання'));
  body.push(p('Алгоритм виконується за q ітерацій (k = 0, 1, ..., q-1):'));
  body.push(bullet('Broadcast: процес (i, k mod q) розсилає свій блок A усім процесам рядка i: (i,0), (i,1), ..., (i,q-1). Операція: MPI_Bcast у горизонтальному комунікаторі.'));
  body.push(bullet('Multiply: кожен процес (i,j) обчислює C_ij += A_temp × B_ij, де A_temp — отриманий broadcast-блок.'));
  body.push(bullet('Shift: блоки B зсуваються вгору по стовпцях: процес (i,j) надсилає B_ij процесу ((i-1+q) mod q, j). Операція: MPI_Sendrecv по вертикальному комунікаторі.'));
  
  body.push(h2('12.3. Приклад (q=2, 4 процеси)'));
  body.push(p('Крок 0: (0,0) broadcasts A_00 по рядку 0 → (0,0) обч. C_00 += A_00×B_00, (0,1) обч. C_01 += A_00×B_01. (1,1) broadcasts A_11 по рядку 1 → (1,0) обч. C_10 += A_11×B_10, (1,1) обч. C_11 += A_11×B_11. Потім B зсуваються вгору.'));
  body.push(p('Крок 1: (0,1) broadcasts A_01 по рядку 0. Множення. Зсув B. Після q кроків кожен процес містить фінальний C_ij.'));
  
  body.push(h2('12.4. Складність'));
  body.push(p('Обчислення: T_comp = q × 2(n/q)³ = 2n³/q² = 2n³/p.'));
  body.push(p('Комунікація за крок: broadcast блоку (n/q)² елементів по q процесах: T_bcast = α×log(q) + β×(n/q)². Зсув B: T_shift = α + β×(n/q)². Загальна комунікація за q кроків: O(n²/q × log q).'));
  body.push(p('Прискорення: S ≈ p / (1 + O(log p / n)). При великих n ефективність наближається до 1.'));
  
  body.push(h2('12.5. Реалізація в Java (структура)'));
  body.push(...codes([
    '// Структура Java-реалізації Fox через Thread на спільній пам\'яті',
    'CyclicBarrier barrier = new CyclicBarrier(q * q);',
    '',
    'class FoxWorker extends Thread {',
    '    int row, col; // позиція в сітці',
    '    double[][] localA, localB, localC, tempA;',
    '',
    '    public void run() {',
    '        for (int step = 0; step < q; step++) {',
    '            int broadcastCol = (row + step) % q;',
    '            // Копіюємо A від broadcast-процесу через спільний масив',
    '            if (col == broadcastCol) sharedA[row] = localA;',
    '            try { barrier.await(); } catch (Exception e) {}',
    '            tempA = sharedA[row];',
    '            // Множимо tempA * localB',
    '            multiplyAdd(tempA, localB, localC);',
    '            try { barrier.await(); } catch (Exception e) {}',
    '            // Зсув B вгору (через спільний буфер)',
    '            shiftBUp();',
    '            try { barrier.await(); } catch (Exception e) {}',
    '        }',
    '    }',
    '}',
  ]));
  
  
  // ═══════════════════════════════════════════════════════════════════
  // ПИТАННЯ 14
  // ═══════════════════════════════════════════════════════════════════
  body.push(h1('Питання 14. Призупинка та відновлення роботи потоку. Переривання'));
  
  body.push(h2('14.1. Застарілі методи (deprecated)'));
  body.push(p('Методи Thread.suspend() та Thread.resume() є застарілими (deprecated) і небезпечними: suspend() заморожує потік разом з усіма захопленими моніторами, що спричиняє deadlock. Thread.stop() кидає ThreadDeath в довільний момент, залишаючи об\'єкти в неузгодженому стані. Ніколи не використовуйте їх у нових програмах.'));
  
  body.push(h2('14.2. Коректна призупинка через wait/notify'));
  body.push(p('Рекомендований підхід: кооперативна призупинка через volatile-прапорець та wait()/notify():'));
  body.push(...codes([
    'class PausableWorker extends Thread {',
    '    private volatile boolean paused = false;',
    '    private volatile boolean stopped = false;',
    '    private final Object pauseLock = new Object();',
    '',
    '    public void run() {',
    '        while (!stopped) {',
    '            synchronized (pauseLock) {',
    '                while (paused) {',
    '                    try { pauseLock.wait(); }',
    '                    catch (InterruptedException e) {',
    '                        Thread.currentThread().interrupt();',
    '                        return;',
    '                    }',
    '                }',
    '            }',
    '            doWork(); // корисна робота',
    '        }',
    '    }',
    '',
    '    public void pause() {',
    '        paused = true;',
    '    }',
    '',
    '    public void resumeWork() {',
    '        synchronized (pauseLock) {',
    '            paused = false;',
    '            pauseLock.notifyAll();',
    '        }',
    '    }',
    '',
    '    public void stopWorker() { stopped = true; interrupt(); }',
    '}',
  ]));
  
  body.push(h2('14.3. Механізм переривання (Interruption)'));
  body.push(p('interrupt() встановлює прапорець переривання потоку. Якщо потік знаходиться в стані WAITING або TIMED_WAITING (sleep, wait, join), він прокидається з InterruptedException та прапорець скидається. Якщо потік RUNNABLE — лише встановлюється прапорець, який слід перевіряти.'));
  body.push(...codes([
    'Thread worker = new Thread(() -> {',
    '    while (!Thread.currentThread().isInterrupted()) {',
    '        try {',
    '            process();',
    '            Thread.sleep(100); // кидає InterruptedException якщо interrupted',
    '        } catch (InterruptedException e) {',
    '            // ВАЖЛИВО: відновити прапорець після перехоплення',
    '            Thread.currentThread().interrupt();',
    '            break; // або return',
    '        }',
    '    }',
    '    cleanup(); // ресурси звільняємо тут',
    '});',
    'worker.start();',
    'Thread.sleep(1000);',
    'worker.interrupt(); // коректне завершення',
    'worker.join();',
  ]));
  
  body.push(h2('14.4. sleep() у деталях'));
  body.push(p('Thread.sleep(millis) — статичний метод, призупиняє поточний потік на задану кількість мілісекунд. НЕ звільняє монітори (на відміну від wait()). Точність залежить від ОС (зазвичай ±10 мс на Windows, ~1 мс на Linux). TimeUnit.MILLISECONDS.sleep(100) — зручніший варіант.'));
  
  
  // ═══════════════════════════════════════════════════════════════════
  // ПИТАННЯ 16
  // ═══════════════════════════════════════════════════════════════════
  body.push(h1('Питання 16. Синхронізація в паралельних обчисленнях'));
  
  body.push(h2('16.1. Навіщо потрібна синхронізація'));
  body.push(p('Race condition (гонка даних) — ситуація, коли результат залежить від непередбачуваного порядку виконання операцій кількома потоками. Приклад: count++ насправді є трьома операціями: read, increment, write. Якщо два потоки виконують це одночасно, одне збільшення може бути втрачене.'));
  body.push(...codes([
    'int count = 0;',
    '// Потік 1: read(0) → +1 → write(1)',
    '// Потік 2: read(0) → +1 → write(1)  // читає 0, бо потік 1 ще не записав!',
    '// Результат: 1 замість 2 — race condition!',
    '',
    '// Виправлення:',
    'AtomicInteger atomicCount = new AtomicInteger(0);',
    'atomicCount.incrementAndGet(); // атомарна операція CAS',
  ]));
  
  body.push(h2('16.2. Засоби синхронізації в Java'));
  body.push(bullet('synchronized — ключове слово: mutex на рівні JVM. Найпростіший, але найповільніший варіант.'));
  body.push(bullet('volatile — видимість (не атомарність) змін між потоками. Без кешування в регістрах.'));
  body.push(bullet('java.util.concurrent.atomic — AtomicInteger, AtomicLong, AtomicReference: lock-free операції через CAS (Compare-And-Swap).'));
  body.push(bullet('ReentrantLock — явне блокування з tryLock, fairness, умовними змінними (Condition).'));
  body.push(bullet('ReadWriteLock — паралельне читання, ексклюзивний запис.'));
  body.push(bullet('StampedLock (Java 8) — оптимістичне читання без захоплення lock.'));
  body.push(bullet('Semaphore — контроль кількості одночасних доступів.'));
  body.push(bullet('CountDownLatch, CyclicBarrier, Phaser — бар\'єрна синхронізація.'));
  
  body.push(h2('16.3. Паттерн Producer-Consumer'));
  body.push(p('Класичний паттерн синхронізації: виробник генерує елементи, споживач їх обробляє. Між ними — буфер з обмеженою ємністю. У Java: BlockingQueue (LinkedBlockingQueue, ArrayBlockingQueue) автоматизує синхронізацію:'));
  body.push(...codes([
    'BlockingQueue<Task> queue = new ArrayBlockingQueue<>(100);',
    '',
    '// Виробник',
    'Executors.newSingleThreadExecutor().submit(() -> {',
    '    for (int i = 0; i < 1000; i++) {',
    '        queue.put(new Task(i)); // блокує якщо черга повна',
    '    }',
    '});',
    '',
    '// Споживач',
    'Executors.newFixedThreadPool(4).submit(() -> {',
    '    while (true) {',
    '        Task t = queue.take(); // блокує якщо черга порожня',
    '        process(t);',
    '    }',
    '});',
  ]));
  
  body.push(h2('16.4. MPI-синхронізація'));
  body.push(p('В розподілених системах синхронізація через MPI_Barrier(comm) — всі процеси чекають найповільнішого. Колективні операції (MPI_Reduce, MPI_Gather) мають вбудовану неявну синхронізацію. Надмірні бар\'єри знижують ефективність.'));
  
  
  // ═══════════════════════════════════════════════════════════════════
  // ПИТАННЯ 18
  // ═══════════════════════════════════════════════════════════════════
  body.push(h1('Питання 18. Блокування об\'єкту'));
  
  body.push(h2('18.1. Монітор об\'єкта Java'));
  body.push(p('Кожен Java-об\'єкт асоційований з монітором (mutex). Монітор — це механізм взаємного виключення: лише один потік одночасно може виконувати synchronized-код на даному об\'єкті. Монітор реалізований у JVM через ObjectMonitor у C++ (HotSpot). Стан монітора зберігається в object header (mark word).'));
  
  body.push(h2('18.2. Синхронізований метод та блок'));
  body.push(...codes([
    'class Account {',
    '    private double balance;',
    '',
    '    // Монітор = this (об\'єкт Account)',
    '    public synchronized void deposit(double amount) {',
    '        balance += amount;',
    '    }',
    '',
    '    // Монітор = Account.class (статичні методи)',
    '    public static synchronized void classMethod() { ... }',
    '',
    '    // Явний вибір монітора — дрібніша гранулярність',
    '    private final Object lock = new Object();',
    '    public void withdraw(double amount) {',
    '        synchronized (lock) {',
    '            if (balance >= amount) balance -= amount;',
    '        }',
    '    }',
    '}',
  ]));
  
  body.push(h2('18.3. Biased Locking, Thin Lock, Inflated Lock'));
  body.push(p('JVM оптимізує блокування поетапно:'));
  body.push(bullet('Biased Locking: якщо об\'єкт завжди захоплюється одним потоком, JVM записує ID потоку у mark word — повторний захват без атомарної операції.'));
  body.push(bullet('Thin Lock (Lightweight Lock): CAS-операція на mark word. Немає ОС-блокування. Швидко, якщо конкуренція низька.'));
  body.push(bullet('Inflated Lock (Heavyweight): якщо thin lock не вдалося — потік блокується через ОС (mutex). Дорого: перемикання контексту.'));
  body.push(p('Java 15+ вимикає Biased Locking як застарілий (JEP 374).'));
  
  body.push(h2('18.4. ReentrantLock як альтернатива'));
  body.push(...codes([
    'ReentrantLock lock = new ReentrantLock(true); // fair=true',
    '',
    '// Умовні змінні — потужніший аналог wait/notify',
    'Condition notFull  = lock.newCondition();',
    'Condition notEmpty = lock.newCondition();',
    '',
    'lock.lock();',
    'try {',
    '    while (buffer.isFull()) notFull.await();',
    '    buffer.add(item);',
    '    notEmpty.signal(); // будимо тільки consumer, не всіх',
    '} finally {',
    '    lock.unlock(); // обов\'язково у finally!',
    '}',
  ]));
  
  body.push(h2('18.5. Атомарні змінні без блокувань'));
  body.push(p('CAS (Compare-And-Swap) — апаратна інструкція: якщо поточне значення == expected, атомарно встановити new. Lock-free алгоритми на основі CAS уникають OS-блокувань:'));
  body.push(...codes([
    'AtomicInteger counter = new AtomicInteger(0);',
    '// Безпечний increment без synchronized:',
    'counter.incrementAndGet(); // CAS-loop всередині',
    '',
    '// Ручний CAS:',
    'int old = counter.get();',
    'while (!counter.compareAndSet(old, old + 1)) {',
    '    old = counter.get(); // retry якщо інший потік змінив',
    '}',
  ]));
  
  
  // ═══════════════════════════════════════════════════════════════════
  // ПИТАННЯ 20
  // ═══════════════════════════════════════════════════════════════════
  body.push(h1('Питання 20. Проблеми управління потоками: дедлок та інші небажані стани'));
  
  body.push(h2('20.1. Deadlock (взаємне блокування)'));
  body.push(p('Deadlock — ситуація, коли два або більше потоків взаємно чекають ресурси, що утримуються один одним. Умови Коффмана (всі чотири мають виконуватись одночасно):'));
  body.push(bullet('Взаємне виключення: ресурс може утримуватись лише одним потоком.'));
  body.push(bullet('Утримання та очікування: потік утримує ресурс і чекає на інший.'));
  body.push(bullet('Відсутність примусового вивільнення: ресурс вивільняється лише добровільно.'));
  body.push(bullet('Кругове очікування: потік 1 чекає потік 2, потік 2 чекає потік 1.'));
  body.push(...codes([
    '// Класичний deadlock',
    'Object lockA = new Object(), lockB = new Object();',
    '',
    'Thread t1 = new Thread(() -> {',
    '    synchronized(lockA) {',
    '        Thread.sleep(10);',
    '        synchronized(lockB) { /* ... */ } // чекає lockB',
    '    }',
    '});',
    'Thread t2 = new Thread(() -> {',
    '    synchronized(lockB) {',
    '        Thread.sleep(10);',
    '        synchronized(lockA) { /* ... */ } // чекає lockA',
    '    }',
    '});',
    '// Обидва потоки заблоковані назавжди!',
    '',
    '// Рішення: завжди захоплювати у фіксованому порядку',
    '// або використовувати tryLock з таймаутом',
  ]));
  
  body.push(h2('20.2. Livelock'));
  body.push(p('Livelock — потоки не заблоковані, але постійно змінюють стан у відповідь один на одного, без прогресу. Аналогія: два люди в коридорі, що постійно поступаються один одному в той самий бік. Рішення: випадкова затримка перед повторною спробою.'));
  
  body.push(h2('20.3. Starvation (голодування)'));
  body.push(p('Starvation — потік ніколи не отримує CPU-час через несправедливість планувальника або постійну конкуренцію з вищопріоритетними потоками. Рішення: fair locks (ReentrantLock(true)), уникати надмірних пріоритетів.'));
  
  body.push(h2('20.4. Priority Inversion'));
  body.push(p('Інверсія пріоритетів: низькопріоритетний потік L утримує монітор, що потрібен високопріоритетному H. Середньопріоритетний M витісняє L, тому H чекає, поки M завершиться. Фактично H має пріоритет нижче M. Рішення: Priority Inheritance Protocol (в RTOS).'));
  
  body.push(h2('20.5. Діагностика'));
  body.push(p('jstack <pid> — виводить стан усіх потоків JVM, включаючи deadlock-аналіз. Java Flight Recorder — для livelock та starvation. ThreadMXBean.findDeadlockedThreads() — програмне виявлення deadlock у runtime.'));
  
  
  // ═══════════════════════════════════════════════════════════════════
  // ПИТАННЯ 22
  // ═══════════════════════════════════════════════════════════════════
  body.push(h1('Питання 22. Локери та управління потоками'));
  
  body.push(h2('22.1. Інтерфейс Lock'));
  body.push(p('java.util.concurrent.locks.Lock — гнучкіша альтернатива synchronized. Надає: явне захоплення/звільнення, tryLock() (без блокування або з таймаутом), lockInterruptibly() (переривання при очікуванні), підтримку кількох Condition об\'єктів.'));
  body.push(...codes([
    'Lock lock = new ReentrantLock();',
    '',
    '// Базове використання',
    'lock.lock();',
    'try { criticalSection(); }',
    'finally { lock.unlock(); }', // завжди у finally!
    '',
    '// Спроба без блокування',
    'if (lock.tryLock()) {',
    '    try { criticalSection(); }',
    '    finally { lock.unlock(); }',
    '} else {',
    '    // інша логіка якщо lock зайнятий',
    '}',
    '',
    '// З таймаутом',
    'if (lock.tryLock(200, TimeUnit.MILLISECONDS)) { ... }',
  ]));
  
  body.push(h2('22.2. ReentrantLock'));
  body.push(p('Найпоширеніша реалізація Lock. Реентрантний: потік може захопити lock, який вже утримує, без deadlock (лічильник входжень). Параметр fair=true забезпечує FIFO-порядок очікуючих потоків (запобігає starvation), але знижує throughput.'));
  
  body.push(h2('22.3. ReadWriteLock та StampedLock'));
  body.push(...codes([
    '// ReadWriteLock: паралельне читання, ексклюзивний запис',
    'ReadWriteLock rwLock = new ReentrantReadWriteLock();',
    'Lock readLock  = rwLock.readLock();',
    'Lock writeLock = rwLock.writeLock();',
    '',
    '// StampedLock (Java 8): оптимістичне читання',
    'StampedLock sl = new StampedLock();',
    'long stamp = sl.tryOptimisticRead(); // без захоплення lock',
    'double x = data.x; double y = data.y;',
    'if (!sl.validate(stamp)) {          // перевірка — чи не було запису?',
    '    stamp = sl.readLock();          // fallback до звичайного readLock',
    '    try { x = data.x; y = data.y; }',
    '    finally { sl.unlockRead(stamp); }',
    '}',
  ]));
  
  body.push(h2('22.4. Condition (умовні змінні)'));
  body.push(p('Condition — аналог Object.wait()/notify(), але асоційований з конкретним Lock. Дозволяє мати кілька незалежних умов очікування на одному Lock, що важливо для паттерну Producer-Consumer з розділеними умовами "не порожньо" та "не повно".'));
  
  body.push(h2('22.5. LockSupport'));
  body.push(p('Низькорівневий примітив: LockSupport.park() — блокує поточний потік. LockSupport.unpark(thread) — розблоковує конкретний потік. Основа реалізації AbstractQueuedSynchronizer (AQS), на якому побудовані всі lock-класи JUC.'));
  
  
  // ═══════════════════════════════════════════════════════════════════
  // ПИТАННЯ 24
  // ═══════════════════════════════════════════════════════════════════
  body.push(h1('Питання 24. Інтерфейс Executor та бібліотечні класи'));
  
  body.push(h2('24.1. Ієрархія Executor Framework'));
  body.push(p('Executor (execute(Runnable)) <- ExecutorService (submit, shutdown, invokeAll, invokeAny) <- ScheduledExecutorService (schedule, scheduleAtFixedRate). Основна реалізація: ThreadPoolExecutor. Фасадний клас: Executors з фабричними методами.'));
  
  body.push(h2('24.2. Executor та ExecutorService'));
  body.push(...codes([
    '// Executor — лише запуск задачі',
    'Executor simple = command -> new Thread(command).start();',
    'simple.execute(() -> System.out.println("Hello"));',
    '',
    '// ExecutorService — повне управління',
    'ExecutorService es = Executors.newFixedThreadPool(4);',
    '',
    '// submit Runnable — Future<?>',
    'Future<?> f1 = es.submit(() -> doWork());',
    '',
    '// submit Callable — Future<V>',
    'Future<Integer> f2 = es.submit(() -> compute());',
    'Integer result = f2.get(); // блокує до завершення',
    'Integer result2 = f2.get(5, TimeUnit.SECONDS); // з таймаутом',
    '',
    '// Пакетне виконання',
    'List<Callable<Integer>> tasks = List.of(() -> 1, () -> 2, () -> 3);',
    'List<Future<Integer>> futures = es.invokeAll(tasks);',
    '',
    '// Завершення',
    'es.shutdown();                    // нові задачі не приймаються',
    'es.awaitTermination(10, TimeUnit.SECONDS); // чекаємо завершення',
  ]));
  
  body.push(h2('24.3. Фабричні методи Executors'));
  body.push(bullet('newFixedThreadPool(n): n потоків, unbounded LinkedBlockingQueue. Для CPU-bound задач n ≈ кількість ядер.'));
  body.push(bullet('newCachedThreadPool(): потоки створюються за потребою, видаляються через 60 с. Для I/O-bound або короткочасних задач.'));
  body.push(bullet('newSingleThreadExecutor(): один потік, задачі виконуються послідовно у порядку FIFO. Гарантує порядок.'));
  body.push(bullet('newScheduledThreadPool(n): для задач за розкладом. scheduleAtFixedRate(task, 0, 1, SECONDS).'));
  body.push(bullet('newWorkStealingPool(): ForkJoinPool з parallelism = Runtime.availableProcessors(). Work-stealing.'));
  
  body.push(h2('24.4. ScheduledExecutorService'));
  body.push(...codes([
    'ScheduledExecutorService scheduler = Executors.newScheduledThreadPool(2);',
    '',
    '// Одноразово через 5 секунд',
    'scheduler.schedule(() -> doTask(), 5, TimeUnit.SECONDS);',
    '',
    '// Кожні 2 секунди (fixed rate)',
    'scheduler.scheduleAtFixedRate(() -> monitor(), 0, 2, TimeUnit.SECONDS);',
    '',
    '// З паузою 2 секунди між завершенням та початком',
    'scheduler.scheduleWithFixedDelay(() -> poll(), 0, 2, TimeUnit.SECONDS);',
  ]));
  
  body.push(h2('24.5. CompletionService'));
  body.push(p('ExecutorCompletionService дозволяє отримувати результати задач у порядку їх завершення (не подання). Корисно, коли потрібно обробити першу готову задачу:'));
  body.push(...codes([
    'CompletionService<String> cs = new ExecutorCompletionService<>(executor);',
    'for (Callable<String> t : tasks) cs.submit(t);',
    'for (int i = 0; i < tasks.size(); i++) {',
    '    Future<String> done = cs.take(); // бере першу завершену',
    '    process(done.get());',
    '}',
  ]));
  
  
  // ═══════════════════════════════════════════════════════════════════
  // ПИТАННЯ 26
  // ═══════════════════════════════════════════════════════════════════
  body.push(h1('Питання 26. Методи моделювання паралельних обчислень'));
  
  body.push(h2('26.1. Навіщо моделювати'));
  body.push(p('Паралельні програми важко аналізувати через недетермінізм (порядок виконання потоків непередбачуваний). Моделювання дозволяє: верифікувати коректність (відсутність deadlock, race), оцінити продуктивність до реалізації, документувати поведінку системи.'));
  
  body.push(h2('26.2. Мережі Петрі'));
  body.push(p('Мережа Петрі (Petri Net) — математична модель дискретних систем. Елементи: позиції (places — стани), переходи (transitions — події/дії), дуги (arcs — зв\'язки), маркери (tokens — ресурси/стан). Правило спрацьовування: перехід спрацьовує, якщо у всіх вхідних позиціях є маркери. Маркери споживаються та породжуються.'));
  body.push(p('Застосування в паралельних програмах: моделювання mutex (маркер = доступний lock), виробника-споживача (маркери в буфері), бар\'єрів (перехід чекає маркерів від усіх потоків).'));
  
  body.push(h2('26.3. Граф "операції-операнди" (DAG)'));
  body.push(p('DAG (Directed Acyclic Graph) відображає залежності між обчисленнями. Вузли — операції або дані, ребра — потоки даних. Вузли без вхідних ребер (initial) можуть виконуватись одразу. Паралелізм виявляється автоматично: вузли без взаємних залежностей — паралельні.'));
  body.push(p('Критичний шлях у DAG визначає нижню межу часу виконання: T_min = довжина найдовшого шляху від входу до виходу.'));
  
  body.push(h2('26.4. Process Calculi'));
  body.push(p('CSP (Communicating Sequential Processes, Хоар) та CCS (Calculus of Communicating Systems, Мілнер) — формальні мови опису паралельних систем. В Java: бібліотека JCSP реалізує CSP. Канали — примітив комунікації. Дозволяють верифікувати відсутність deadlock формально.'));
  
  body.push(h2('26.5. Симуляція та профілювання'));
  body.push(p('Java Pathfinder (JPF) — model checker для Java: перебирає всі можливі порядки виконання потоків, виявляє deadlock та race conditions. JProfiler, async-profiler — практичне профілювання для виявлення вузьких місць. ThreadSanitizer (TSan) — для C/C++/MPI-програм.'));
  
  
  // ═══════════════════════════════════════════════════════════════════
  // ПИТАННЯ 28
  // ═══════════════════════════════════════════════════════════════════
  body.push(h1('Питання 28. Моделювання паралельних програм мережею Петрі'));
  
  body.push(h2('28.1. Формальне визначення'));
  body.push(p('Мережа Петрі визначається як п\'ятірка PN = (P, T, F, W, M0), де:'));
  body.push(bullet('P — скінченна множина позицій (кола на діаграмі).'));
  body.push(bullet('T — скінченна множина переходів (прямокутники), P ∩ T = Ø.'));
  body.push(bullet('F ⊆ (P×T) ∪ (T×P) — множина дуг.'));
  body.push(bullet('W: F → N — вага дуг (скільки маркерів споживається/породжується).'));
  body.push(bullet('M0: P → N — початкове маркування (початкові маркери у позиціях).'));
  
  body.push(h2('28.2. Правило спрацьовування'));
  body.push(p('Перехід t ввімкнений (enabled) якщо для кожної вхідної позиції p: M(p) >= W(p,t). Спрацьовування: M\'(p) = M(p) - W(p,t) для вхідних позицій; M\'(p) = M(p) + W(t,p) для вихідних.'));
  
  body.push(h2('28.3. Моделювання mutex'));
  body.push(p('Моделюємо mutex для двох потоків. Позиції: P_idle1 (потік 1 вільний), P_idle2 (потік 2 вільний), P_mutex (маркер = lock вільний), P_cs1 (потік 1 у критичній секції), P_cs2. Переходи: T_enter1 споживає P_idle1 та P_mutex, породжує P_cs1. T_exit1: споживає P_cs1, породжує P_idle1 та P_mutex. Взаємне виключення: лише один маркер у P_mutex гарантує, що T_enter1 та T_enter2 не можуть спрацювати одночасно.'));
  
  body.push(h2('28.4. Моделювання Producer-Consumer'));
  body.push(p('Позиції: P_prod (виробник готовий), P_buffer (маркери = кількість елементів у буфері, max N), P_empty (маркери = вільні місця = N - заповнено), P_cons (споживач готовий). Переходи: T_produce: P_prod + P_empty → P_buffer + P_prod. T_consume: P_buffer + P_cons → P_cons. Deadlock: якщо P_buffer=0 і потік-споживач чекає — маркер у P_buffer відновиться від виробника.'));
  
  body.push(h2('28.5. Властивості мереж Петрі'));
  body.push(bullet('Достижимість (Reachability): чи можна досягти маркування M з M0.'));
  body.push(bullet('Безпечність (Safety): кількість маркерів у кожній позиції обмежена.'));
  body.push(bullet('Живість (Liveness): кожен перехід може спрацювати в деякому досяжному маркуванні. Якщо мережа не жива — є deadlock.'));
  body.push(bullet('Збереженість (Conservation): сума маркерів не змінюється (для консервативних мереж).'));
  
  body.push(h2('28.6. Розширення: Кольорові мережі Петрі'));
  body.push(p('Colored Petri Nets (CPN) — маркери мають тип/колір (клас даних). Один маркер = екземпляр даних з атрибутами. Дозволяє моделювати конкретні значення, не лише присутність/відсутність. Інструмент: CPN Tools (безкоштовний).'));
  
  
  // ═══════════════════════════════════════════════════════════════════
  // ПИТАННЯ 30
  // ═══════════════════════════════════════════════════════════════════
  body.push(h1('Питання 30. Алгоритми паралельного сумування та оцінка їх ефективності'));
  
  body.push(h2('30.1. Послідовне сумування'));
  body.push(p('Задача: обчислити S = Σ a[i] (i=0..n-1). Послідовна складність: O(n), T1 = n операцій. Стратегія паралелізації: розбиваємо масив на частини, кожна частина — локальна сума, потім об\'єднуємо.'));
  
  body.push(h2('30.2. Лінійне розбиття (Flat Decomposition)'));
  body.push(p('p потоків, кожен обчислює локальну суму n/p елементів. Потім головний потік складає p локальних сум. T(p) = n/p + p. Оптимальне p = √n, тоді T = 2√n. Прискорення S = n / 2√n = √n / 2.'));
  body.push(...codes([
    'int chunk = n / numThreads;',
    'double[] localSums = new double[numThreads];',
    'Thread[] threads = new Thread[numThreads];',
    '',
    'for (int t = 0; t < numThreads; t++) {',
    '    final int id = t;',
    '    final int from = id * chunk;',
    '    final int to = (id == numThreads-1) ? n : from + chunk;',
    '    threads[t] = new Thread(() -> {',
    '        double sum = 0;',
    '        for (int i = from; i < to; i++) sum += a[i];',
    '        localSums[id] = sum;',
    '    });',
    '    threads[t].start();',
    '}',
    'for (Thread th : threads) th.join();',
    'double total = Arrays.stream(localSums).sum();',
  ]));
  
  body.push(h2('30.3. Деревоподібне (двійкове) редукування'));
  body.push(p('Алгоритм: на кроці k, потоки з парними індексами складають свій результат з результатом сусіда (парний i додає непарний i+1). Кількість кроків: log₂(p). T(p) = n/p + log₂(p).'));
  body.push(p('Прискорення: S = n / (n/p + log₂(p)) ≈ p при n >> p×log(p). Ефективність: E = S/p = 1/(1 + p×log(p)/n). При p = 4, n = 1000: E ≈ 0.98.'));
  body.push(...codes([
    '// Деревоподібна редукція з CyclicBarrier',
    'CyclicBarrier barrier = new CyclicBarrier(p);',
    'double[] partial = new double[p]; // заповнено локальними сумами',
    '',
    'for (int step = 1; step < p; step *= 2) {',
    '    // Кожен потік t, де t % (2*step) == 0, складає з t+step',
    '    if (myId % (2 * step) == 0 && myId + step < p)',
    '        partial[myId] += partial[myId + step];',
    '    barrier.await();',
    '}',
    '// partial[0] — загальна сума',
  ]));
  
  body.push(h2('30.4. Паралельне сумування в MPI'));
  body.push(p('Найпростіше: MPI_Reduce(localSum, globalSum, 1, MPI_DOUBLE, MPI_SUM, 0, comm). Внутрішньо реалізує деревоподібну редукцію. Час: T_reduce = α×log(p) + β×log(p). MPI_Allreduce — результат отримують всі процеси (для broadcast суми).'));
  
  body.push(h2('30.5. parallelStream у Java'));
  body.push(...codes([
    'double sum = IntStream.range(0, n)',
    '    .parallel()',
    '    .mapToDouble(i -> a[i])',
    '    .sum(); // автоматична деревоподібна редукція',
  ]));
  
  
  // ═══════════════════════════════════════════════════════════════════
  // ПИТАННЯ 32
  // ═══════════════════════════════════════════════════════════════════
  body.push(h1('Питання 32. Експериментальне дослідження ефективності паралельних обчислень'));
  
  body.push(h2('32.1. Методологія бенчмаркінгу'));
  body.push(p('Коректний бенчмарк паралельного алгоритму має враховувати: warmup JIT (JVM оптимізує код після ~10000 ітерацій), garbage collection, NUMA-ефекти, hyperthreading, cache warming. Стандарт для Java: JMH (Java Microbenchmark Harness).'));
  body.push(...codes([
    '// Структура JMH-бенчмарку',
    '@BenchmarkMode(Mode.AverageTime)',
    '@OutputTimeUnit(TimeUnit.MILLISECONDS)',
    '@Warmup(iterations = 5, time = 1)',
    '@Measurement(iterations = 10, time = 1)',
    '@Fork(2)',
    'public class MatrixBenchmark {',
    '    @Benchmark',
    '    public void sequential(Blackhole bh) { bh.consume(seqMultiply()); }',
    '',
    '    @Benchmark',
    '    public void parallel4(Blackhole bh) { bh.consume(parMultiply(4)); }',
    '}',
  ]));
  
  body.push(h2('32.2. Що вимірюємо'));
  body.push(bullet('Wall-clock time (System.nanoTime()) — реальний час від початку до кінця. Включає I/O, GC, OS.'));
  body.push(bullet('CPU time (ThreadMXBean.getCurrentThreadCpuTime()) — час CPU, що витратив конкретний потік.'));
  body.push(bullet('Throughput — кількість задач на секунду.'));
  body.push(bullet('Latency — час виконання однієї задачі.'));
  body.push(bullet('Scalability — вимірювання для p = 1, 2, 4, 8, 16, 32 процесорів.'));
  
  body.push(h2('32.3. Побудова графіків ефективності'));
  body.push(p('Графік "прискорення від кількості потоків": вісь X — кількість потоків (1, 2, 4, 8, 16, 32), вісь Y — S(p) = T(1)/T(p). Ідеальна лінія: S=p. Реальна крива насичується (закон Амдала). Порівнюємо з теоретичним Амдалом: S = 1/(f + (1-f)/p), знаходимо f.'));
  body.push(p('Графік "ефективність": E(p) = S(p)/p. Спадна крива: чим більше потоків, тим нижча ефективність кожного. Ціль: E > 0.7 для практичної доцільності.'));
  
  body.push(h2('32.4. Джерела похибок'));
  body.push(bullet('JIT-компіляція: перші запуски повільніші. Рішення: warmup-ітерації.'));
  body.push(bullet('GC: непередбачувані паузи. Рішення: G1/ZGC, примусовий GC перед вимірюванням.'));
  body.push(bullet('Variance: розкид результатів. Рішення: багато ітерацій, confidence interval.'));
  body.push(bullet('Dead code elimination: JVM може не виконувати код без side effects. Рішення: Blackhole в JMH.'));
  
  body.push(h2('32.5. Практичний висновок'));
  body.push(p('Для матриць 1000×1000 на 8-ядерному CPU: очікувана S(8) ≈ 6–7.5 (залежно від алгоритму). False sharing може знизити S(8) до 3–4. Cache-friendly (рядково-орієнтоване) зберігання матриць критично для cache hit rate.'));
  
  
  // ═══════════════════════════════════════════════════════════════════
  // ПИТАННЯ 34
  // ═══════════════════════════════════════════════════════════════════
  body.push(h1('Питання 34. Проектування паралельних програм'));
  
  body.push(h2('34.1. Методологія PCAM'));
  body.push(p('Систематичний підхід до проектування паралельних програм — методологія Фостера PCAM (Partition, Communicate, Agglomerate, Map):'));
  body.push(bullet('Partition (Розбиття): розбити задачу та дані на максимально дрібні незалежні частини. Два підходи: Domain Decomposition (розбиття даних) та Functional Decomposition (розбиття обчислень).'));
  body.push(bullet('Communicate (Комунікація): визначити необхідні обміни даними між частинами. Мінімізувати обсяг і кількість комунікацій. Локальні комунікації кращі за глобальні.'));
  body.push(bullet('Agglomerate (Агломерація): об\'єднати дрібні задачі в більші для зменшення overhead комунікацій. Знайти баланс між паралелізмом та granularity.'));
  body.push(bullet('Map (Відображення): призначити агломеровані задачі конкретним процесорам/потокам. Мета: мінімізація комунікацій між вузлами + балансування навантаження.'));
  
  body.push(h2('34.2. Domain Decomposition'));
  body.push(p('Для задач з великими масивами даних: ділимо масив/матрицю між процесами. Підходи: рядковий (row), стовпцевий (column), блочний (block), рекурсивний (quad-tree). Вибір залежить від паттерну доступу алгоритму.'));
  
  body.push(h2('34.3. Балансування навантаження'));
  body.push(p('Статичне балансування: рівномірний розподіл на початку. Ефективне для регулярних задач (матриці). Динамічне: задачі виконуються з черги (work queue), потоки беруть нові задачі після завершення. Ефективне для нерегулярних задач (дерева, графи).'));
  body.push(...codes([
    '// Динамічне балансування через work queue',
    'BlockingQueue<Integer> workQueue = new LinkedBlockingQueue<>();',
    'for (int i = 0; i < n; i++) workQueue.add(i); // n задач',
    '',
    'int P = Runtime.getRuntime().availableProcessors();',
    'ExecutorService pool = Executors.newFixedThreadPool(P);',
    'for (int t = 0; t < P; t++) {',
    '    pool.submit(() -> {',
    '        Integer task;',
    '        while ((task = workQueue.poll()) != null)',
    '            processRow(task); // кожен потік бере рядок',
    '    });',
    '}',
  ]));
  
  body.push(h2('34.4. Проектування синхронізації'));
  body.push(p('Принципи: мінімізувати критичні секції (виносити обчислення за synchronized), уникати вкладених locks (deadlock), надавати перевагу lock-free структурам (AtomicInteger, ConcurrentHashMap), використовувати незмінні (immutable) об\'єкти де можливо.'));
  
  body.push(h2('34.5. Паттерни паралельного програмування'));
  body.push(bullet('Pipeline: обробка потоку даних через послідовність паралельних стадій.'));
  body.push(bullet('Master-Worker: головний процес розподіляє задачі між робочими.'));
  body.push(bullet('MapReduce: map-фаза паралельна, reduce — агрегація результатів.'));
  body.push(bullet('Divide and Conquer: рекурсивне розбиття (ForkJoin).'));
  body.push(bullet('Stencil: кожен елемент залежить від сусідів (чисельні методи).'));
  
  
  // ═══════════════════════════════════════════════════════════════════
  // ПИТАННЯ 36
  // ═══════════════════════════════════════════════════════════════════
  body.push(h1('Питання 36. Моделі пам\'яті паралельних обчислень'));
  
  body.push(h2('36.1. Sequentially Consistent (SC) модель'));
  body.push(p('Сильна модель: результат паралельного виконання еквівалентний деякому послідовному виконанню, де операції кожного потоку зберігають свій порядок. Легко міркувати, але важко ефективно реалізувати на сучасних CPU (вимагає барєрів пам\'яті на кожній операції).'));
  
  body.push(h2('36.2. Total Store Order (TSO)'));
  body.push(p('Модель x86/x64: записи буферизуються у write buffer потоку. Читання бачить своє останнє записане значення, але може не бачити запис іншого потоку одразу. Сильніша за Java MM, але слабша за SC. Пояснює, чому volatile важливий на x86.'));
  
  body.push(h2('36.3. Java Memory Model (JMM)'));
  body.push(p('JMM (JSR-133, Java 5+) визначає happens-before гарантії. Не прив\'язана до конкретної апаратури. Основні правила:'));
  body.push(bullet('Program order rule: дії у потоці впорядковані happens-before відповідно до програмного порядку.'));
  body.push(bullet('Monitor lock rule: unlock(m) happens-before lock(m) будь-яким потоком.'));
  body.push(bullet('Volatile rule: запис у volatile x happens-before читання volatile x.'));
  body.push(bullet('Thread start rule: Thread.start() happens-before будь-яка дія у запущеному потоці.'));
  body.push(bullet('Thread join rule: будь-яка дія у потоці happens-before Thread.join() на цьому потоці.'));
  
  body.push(h2('36.4. Слабкі моделі та Memory Fence'));
  body.push(p('ARM та RISC-V мають слабкіші моделі: читання і записи можуть переупорядковуватись. Memory barrier (fence) — явна інструкція, що запобігає переупорядкуванню. В Java: VarHandle.fullFence(), або synchronized/volatile автоматично вставляють потрібні бар\'єри.'));
  
  body.push(h2('36.5. PRAM-модель'));
  body.push(p('PRAM (Parallel Random Access Machine) — теоретична модель: p процесорів з синхронним доступом до спільної пам\'яті. Варіанти: EREW (Exclusive Read Exclusive Write), CREW (Concurrent Read Exclusive Write), CRCW (Concurrent Read Concurrent Write). Використовується для теоретичного аналізу алгоритмів, не відображає реальне апаратне забезпечення.'));
  
  
  // ═══════════════════════════════════════════════════════════════════
  // ПИТАННЯ 38
  // ═══════════════════════════════════════════════════════════════════
  body.push(h1('Питання 38. Стандарт Message Passing Interface (MPI)'));
  
  body.push(h2('38.1. Що таке MPI'));
  body.push(p('MPI (Message Passing Interface) — стандарт бібліотечного API для передачі повідомлень у паралельних та розподілених обчисленнях. Версії: MPI-1 (1994), MPI-2 (1997, односторонні комунікації), MPI-3 (2012, колективні неблокуючі). Реалізації: OpenMPI, MPICH, Intel MPI, MS MPI. Мови: C, C++, Fortran (та обгортки для Java, Python).'));
  
  body.push(h2('38.2. Базова структура програми'));
  body.push(...codes([
    '#include <mpi.h>',
    'int main(int argc, char** argv) {',
    '    MPI_Init(&argc, &argv);',
    '',
    '    int rank, size;',
    '    MPI_Comm_rank(MPI_COMM_WORLD, &rank); // 0 .. size-1',
    '    MPI_Comm_size(MPI_COMM_WORLD, &size);',
    '',
    '    // Кожен процес виконує свою частину роботи',
    '    double localResult = compute(rank, size);',
    '',
    '    // Агрегація результатів',
    '    double globalResult;',
    '    MPI_Reduce(&localResult, &globalResult, 1,',
    '               MPI_DOUBLE, MPI_SUM, 0, MPI_COMM_WORLD);',
    '',
    '    if (rank == 0)',
    '        printf("Total: %f\\n", globalResult);',
    '',
    '    MPI_Finalize();',
    '}',
    '// Компіляція: mpicc -o prog prog.c',
    '// Запуск:     mpirun -np 8 ./prog',
  ]));
  
  body.push(h2('38.3. Комунікатори'));
  body.push(p('MPI_COMM_WORLD — стандартний комунікатор, що об\'єднує всі процеси. MPI_COMM_SELF — комунікатор лише з поточним процесом. Користувацькі комунікатори: MPI_Comm_split (розподіл за кольором), MPI_Comm_create. Картезіанські топології: MPI_Cart_create для 2D/3D сіток.'));
  
  body.push(h2('38.4. Типи даних MPI'));
  body.push(p('Базові: MPI_INT, MPI_DOUBLE, MPI_FLOAT, MPI_CHAR, MPI_LONG. Похідні типи для нерегулярних структур:'));
  body.push(...codes([
    '// Тип для стовпця матриці (рядково-мажорне зберігання)',
    'MPI_Datatype col_type;',
    'MPI_Type_vector(n, 1, n, MPI_DOUBLE, &col_type); // n блоків по 1, крок n',
    'MPI_Type_commit(&col_type);',
    '// Тепер можна надіслати стовпець:',
    'MPI_Send(&A[0][j], 1, col_type, dest, tag, comm);',
    'MPI_Type_free(&col_type);',
  ]));
  
  body.push(h2('38.5. Топології та картезіанська сітка'));
  body.push(...codes([
    'int dims[2] = {q, q}; // q×q сітка',
    'int periods[2] = {1, 1}; // тороїдальна (cyclic)',
    'MPI_Comm cart_comm;',
    'MPI_Cart_create(MPI_COMM_WORLD, 2, dims, periods, 1, &cart_comm);',
    '',
    'int coords[2];',
    'MPI_Cart_coords(cart_comm, rank, 2, coords); // row=coords[0], col=coords[1]',
    '',
    '// Сусіди для зсуву',
    'int left, right, up, down;',
    'MPI_Cart_shift(cart_comm, 1, 1, &left, &right);   // горизонтальний зсув',
    'MPI_Cart_shift(cart_comm, 0, 1, &up,   &down);    // вертикальний зсув',
  ]));
  
  
  // ═══════════════════════════════════════════════════════════════════
  // ПИТАННЯ 40
  // ═══════════════════════════════════════════════════════════════════
  body.push(h1('Питання 40. Колективні методи обміну повідомленнями MPI'));
  
  body.push(h2('40.1. Огляд колективних операцій'));
  body.push(p('Колективні операції залучають всі процеси комунікатора одночасно. Переваги перед ручними Send/Recv: оптимізована реалізація (деревоподібна, butterfly, binomial tree), читабельний код, автоматична синхронізація.'));
  
  body.push(h2('40.2. MPI_Bcast — широкомовна розсилка'));
  body.push(...codes([
    '// root=0 надсилає buf всім іншим',
    'double buf[N];',
    'if (rank == 0) fillBuffer(buf);',
    'MPI_Bcast(buf, N, MPI_DOUBLE, 0, MPI_COMM_WORLD);',
    '// Тепер buf однаковий у всіх процесів',
    '// Час: O(α*log(p) + β*N*log(p)) з деревом',
  ]));
  
  body.push(h2('40.3. MPI_Scatter та MPI_Gather'));
  body.push(...codes([
    '// Scatter: root розподіляє різні частини масиву між процесами',
    'double send[N*p]; // лише у root',
    'double recv[N];   // у кожного',
    'MPI_Scatter(send, N, MPI_DOUBLE, recv, N, MPI_DOUBLE, 0, comm);',
    '',
    '// Gather: збирає частини від всіх на root',
    'MPI_Gather(recv, N, MPI_DOUBLE, result, N, MPI_DOUBLE, 0, comm);',
    '',
    '// Scatterv/Gatherv: різні розміри для кожного процесу',
    'int counts[p] = {3, 5, 2, 6}; // нерівномірний розподіл',
    'int displs[p] = {0, 3, 8, 10};',
    'MPI_Scatterv(send, counts, displs, MPI_DOUBLE, recv,',
    '             counts[rank], MPI_DOUBLE, 0, comm);',
  ]));
  
  body.push(h2('40.4. MPI_Reduce та MPI_Allreduce'));
  body.push(...codes([
    '// Reduce: агрегація з операцією на root',
    'double local = computeLocalSum();',
    'double global;',
    'MPI_Reduce(&local, &global, 1, MPI_DOUBLE, MPI_SUM, 0, comm);',
    '// Операції: MPI_SUM, MPI_MAX, MPI_MIN, MPI_PROD, MPI_LAND, MPI_LOR',
    '',
    '// Allreduce: результат у всіх (не лише root)',
    'MPI_Allreduce(&local, &global, 1, MPI_DOUBLE, MPI_SUM, comm);',
    '',
    '// Reduce_scatter: редукція + розсилка частин',
    'MPI_Reduce_scatter(send, recv, counts, MPI_DOUBLE, MPI_SUM, comm);',
  ]));
  
  body.push(h2('40.5. MPI_Allgather та MPI_Alltoall'));
  body.push(...codes([
    '// Allgather: кожен збирає дані від усіх (транспонована Gather)',
    'MPI_Allgather(&local, N, MPI_DOUBLE, allData, N, MPI_DOUBLE, comm);',
    '// allData[i*N .. (i+1)*N-1] — дані від процесу i',
    '',
    '// Alltoall: матрична транспозиція комунікацій',
    '// кожен надсилає різні дані кожному',
    'MPI_Alltoall(send, N, MPI_DOUBLE, recv, N, MPI_DOUBLE, comm);',
  ]));
  
  body.push(h2('40.6. MPI_Barrier'));
  body.push(p('Синхронізаційний бар\'єр: жоден процес не продовжує роботу, поки всі не досягли MPI_Barrier(). Використовується для: розмежування фаз обчислень, коректного вимірювання часу (MPI_Wtime() після бар\'єру). Надмірні бар\'єри знижують продуктивність.'));
  
  
  // ═══════════════════════════════════════════════════════════════════
  // ПИТАННЯ 42
  // ═══════════════════════════════════════════════════════════════════
  body.push(h1('Питання 42. Розробка ефективних паралельних алгоритмів в OpenMPI'));
  
  body.push(h2('42.1. Принципи ефективності в MPI'));
  body.push(p('Головні фактори продуктивності MPI-програм:'));
  body.push(bullet('Мінімізація кількості повідомлень: кілька малих повідомлень дорожчі за одне велике (latency dominates). Агломерація: пакувати дані в один буфер.'));
  body.push(bullet('Перекриття комунікацій та обчислень: неблокуючі MPI_Isend/MPI_Irecv дозволяють обчислювати, поки дані передаються.'));
  body.push(bullet('Колективні операції замість P2P: MPI_Bcast швидший за цикл MPI_Send.'));
  body.push(bullet('Правильна топологія: MPI_Cart_create для алгоритмів з регулярним паттерном комунікацій.'));
  
  body.push(h2('42.2. Communication Hiding (перекриття)'));
  body.push(...codes([
    '// Спершу надсилаємо граничні елементи (boundary),',
    '// потім обчислюємо внутрішні (interior) — паралельно з передачею',
    '',
    'MPI_Request reqs[4];',
    '// Неблокуюча відправка меж',
    'MPI_Isend(top_row, N, MPI_DOUBLE, up,   0, comm, &reqs[0]);',
    'MPI_Isend(bot_row, N, MPI_DOUBLE, down, 0, comm, &reqs[1]);',
    'MPI_Irecv(ghost_top, N, MPI_DOUBLE, up,   0, comm, &reqs[2]);',
    'MPI_Irecv(ghost_bot, N, MPI_DOUBLE, down, 0, comm, &reqs[3]);',
    '',
    '// Обчислення внутрішніх рядків — поки йде передача',
    'computeInterior(localGrid, 1, localRows-1);',
    '',
    '// Чекаємо завершення обміну',
    'MPI_Waitall(4, reqs, MPI_STATUSES_IGNORE);',
    '',
    '// Тепер обчислюємо граничні рядки',
    'computeBoundary(localGrid, ghostTop, ghostBot);',
  ]));
  
  body.push(h2('42.3. Derived Datatypes для нерегулярних структур'));
  body.push(p('MPI_Type_vector, MPI_Type_struct, MPI_Type_indexed дозволяють надсилати несуміжні дані без копіювання у тимчасовий буфер:'));
  body.push(...codes([
    '// Надсилаємо стовпець матриці без копіювання',
    'MPI_Datatype column_t;',
    'MPI_Type_vector(rows, 1, cols, MPI_DOUBLE, &column_t);',
    'MPI_Type_commit(&column_t);',
    'MPI_Send(&matrix[0][col], 1, column_t, dest, 0, comm);',
    'MPI_Type_free(&column_t);',
  ]));
  
  body.push(h2('42.4. Профілювання MPI-програм'));
  body.push(p('mpirun -np 4 --mca btl_base_verbose 1 — рівень деталізації комунікацій. TAU, Score-P, Scalasca — профілювальники для MPI. Трасування: mpiP (статистика по функціях), Vampir (візуалізація timeline). Типові проблеми: load imbalance (один процес робить більше), MPI_Barrier latency, excessive small messages.'));
  
  
  // ═══════════════════════════════════════════════════════════════════
  // ПИТАННЯ 44
  // ═══════════════════════════════════════════════════════════════════
  body.push(h1('Питання 44. Паралельна реалізація алгоритму МГУА'));
  
  body.push(h2('44.1. Що таке МГУА'));
  body.push(p('МГУА (Метод Групового Урахування Аргументів, GMDH — Group Method of Data Handling) — евристичний алгоритм для індуктивного моделювання та ідентифікації систем, розроблений О.Г. Івахненком (1968). Будує поліноміальні регресійні моделі шляхом еволюційного відбору кращих комбінацій змінних по шарах.'));
  
  body.push(h2('44.2. Структура алгоритму'));
  body.push(p('МГУА — ітеративний багатошаровий алгоритм:'));
  body.push(bullet('Шар 0: вхідні змінні x1, x2, ..., xm.'));
  body.push(bullet('На кожному шарі: перебираємо всі пари (xi, xj) з поточного шару. Для кожної пари будуємо часткову описову функцію (ЧОФ): поліном 2-го ступеня yi,j = a0 + a1*xi + a2*xj + a3*xi*xj + a4*xi² + a5*xj² (МНК для коефіцієнтів).'));
  body.push(bullet('Відбір: обчислюємо критерій якості (зовнішнє доповнення, регулярність) для кожної ЧОФ. Відбираємо F найкращих як входи наступного шару.'));
  body.push(bullet('Зупинка: якщо критерій якості погіршується або досягнуто максимальну кількість шарів.'));
  
  body.push(h2('44.3. Де паралелізм'));
  body.push(p('МГУА природньо паралельний на рівні: (1) незалежне обчислення ЧОФ для кожної пари змінних, (2) незалежне обчислення критеріїв якості. Кількість пар: C(m, 2) = m*(m-1)/2. Для m=100: 4950 незалежних задач МНК.'));
  
  body.push(h2('44.4. Паралельна реалізація на Java'));
  body.push(...codes([
    'ExecutorService pool = Executors.newFixedThreadPool(Runtime',
    '    .getRuntime().availableProcessors());',
    '',
    'List<Future<NodeResult>> futures = new ArrayList<>();',
    '',
    '// Паралельно будуємо всі ЧОФ поточного шару',
    'for (int i = 0; i < vars; i++) {',
    '    for (int j = i + 1; j < vars; j++) {',
    '        final int fi = i, fj = j;',
    '        futures.add(pool.submit(() ->',
    '            buildPDF(data, fi, fj, trainIdx, validIdx)',
    '        ));',
    '    }',
    '}',
    '',
    '// Збираємо результати',
    'List<NodeResult> results = new ArrayList<>();',
    'for (Future<NodeResult> f : futures)',
    '    results.add(f.get());',
    '',
    '// Відбір F найкращих для наступного шару',
    'results.sort(Comparator.comparingDouble(r -> r.criterion));',
    'List<NodeResult> selected = results.subList(0, F);',
  ]));
  
  body.push(h2('44.5. Розпаралелення на MPI'));
  body.push(p('Для дуже великих m: master розподіляє пари між процесами (MPI_Scatter або динамічна черга). Кожен процес обчислює свій підмножину пар. Master збирає результати (MPI_Gather) та відбирає кращі (локальна сортировка + MPI_Reduce з пошуком max).'));
  
  
  // ═══════════════════════════════════════════════════════════════════
  // ПИТАННЯ 46
  // ═══════════════════════════════════════════════════════════════════
  body.push(h1('Питання 46. Технологія Remote Method Invocation (RMI)'));
  
  body.push(h2('46.1. Що таке RMI'));
  body.push(p('Java RMI (Remote Method Invocation) — механізм виклику методів Java-об\'єктів, що знаходяться в іншій JVM (на іншій машині або тому самому комп\'ютері). Забезпечує прозорість: виклик виглядає як локальний, але виконується дистанційно. Основа для розподілених Java-застосунків до появи EJB та REST.'));
  
  body.push(h2('46.2. Архітектура RMI'));
  body.push(bullet('Remote Interface: інтерфейс, що розширює java.rmi.Remote. Всі методи кидають RemoteException.'));
  body.push(bullet('Remote Object (Server): клас, що реалізує Remote Interface, розширює UnicastRemoteObject.'));
  body.push(bullet('Stub: автоматично генерований проксі на клієнті. Серіалізує аргументи, надсилає мережею, десеріалізує відповідь.'));
  body.push(bullet('Skeleton (Java 1.1) / Dispatcher (Java 2+): на сервері отримує запит, викликає реальний метод.'));
  body.push(bullet('RMI Registry: служба імен. Клієнт отримує stub за ім\'ям через Naming.lookup().'));
  
  body.push(h2('46.3. Реалізація'));
  body.push(...codes([
    '// 1. Remote Interface',
    'public interface MatrixService extends Remote {',
    '    double[][] multiply(double[][] A, double[][] B) throws RemoteException;',
    '}',
    '',
    '// 2. Server Implementation',
    'public class MatrixServiceImpl extends UnicastRemoteObject',
    '                               implements MatrixService {',
    '    public MatrixServiceImpl() throws RemoteException {}',
    '',
    '    public double[][] multiply(double[][] A, double[][] B)',
    '                               throws RemoteException {',
    '        return sequentialMultiply(A, B); // або паралельне!',
    '    }',
    '}',
    '',
    '// 3. Server main',
    'MatrixService service = new MatrixServiceImpl();',
    'Registry registry = LocateRegistry.createRegistry(1099);',
    'registry.bind("MatrixService", service);',
    '',
    '// 4. Client',
    'Registry registry = LocateRegistry.getRegistry("server-host", 1099);',
    'MatrixService service = (MatrixService) registry.lookup("MatrixService");',
    'double[][] C = service.multiply(A, B); // прозорий виклик',
  ]));
  
  body.push(h2('46.4. Серіалізація'));
  body.push(p('Аргументи та результати передаються через Java Serialization. Всі передані об\'єкти мають реалізувати Serializable. Великі матриці: overhead серіалізації значний. Рішення: передавати лише фрагменти або використовувати ефективніші формати (protobuf через gRPC).'));
  
  body.push(h2('46.5. RMI vs REST vs gRPC'));
  body.push(p('RMI: тільки Java, автоматична серіалізація, застарілий для нових проектів. REST: universally supported, JSON, stateless, найпростіший. gRPC: типізований, ефективний (protobuf), streaming, підходить для мікросервісів.'));
  
  body.push(h2('46.6. Spring Boot як альтернатива'));
  body.push(p('У сучасних Java-проектах RMI замінений Spring Boot REST або gRPC. Spring Boot + @RestController + @PostMapping — аналог RMI з REST-інтерфейсом. Spring Boot + WebFlux — реактивний (non-blocking) аналог.'));
  
  
  // ═══════════════════════════════════════════════════════════════════
  // ПИТАННЯ 48
  // ═══════════════════════════════════════════════════════════════════
  body.push(h1('Питання 48. Базові складові грід-системи'));
  
  body.push(h2('48.1. Архітектурні рівні Grid'));
  body.push(p('Архітектура Grid-системи (за Яном Фостером, Ian Foster) складається з п\'яти шарів (sandglass model):'));
  body.push(bullet('Fabric Layer (Рівень ресурсів): фізичні ресурси — обчислювальні вузли, кластери, сховища, наукові прилади. Надають локальні інтерфейси управління (PBS, SLURM для кластерів).'));
  body.push(bullet('Connectivity Layer (Рівень зв\'язності): протоколи комунікацій та автентифікації. TCP/IP для передачі даних, GSI (Grid Security Infrastructure) для PKI-автентифікації, SSL/TLS шифрування.'));
  body.push(bullet('Resource Layer (Рівень ресурсів): протоколи доступу до одного ресурсу. GRAM (запуск задач), GridFTP (передача файлів), OGSI/WSRF для Grid Services.'));
  body.push(bullet('Collective Layer (Колективний рівень): сервіси, що координують кілька ресурсів. Брокер ресурсів, каталоги, моніторинг, репозиторії даних.'));
  body.push(bullet('Application Layer (Прикладний рівень): користувацькі застосунки. Наукові симуляції, бізнес-застосунки, портали.'));
  
  body.push(h2('48.2. Ключові компоненти'));
  body.push(bullet('Resource Broker (Брокер ресурсів): отримує запит, знаходить підходящі ресурси у Grid Information Service, розміщує задачу. Приклади: Condor-G, GRAM, DIANE.'));
  body.push(bullet('Grid Information Service (GIS/MDS): реєстр ресурсів та їх стану. LDAP-сервіс. Дозволяє виявляти ресурси динамічно.'));
  body.push(bullet('Data Management: GridFTP (паралельна передача), репліки (LCG File Catalog), сховища (dCache, iRODS).'));
  body.push(bullet('Security (GSI): X.509 сертифікати, proxy certificates для делегування прав, єдиний вхід (SSO) між організаціями.'));
  body.push(bullet('Monitoring: Ganglia, Nagios для моніторингу стану вузлів. GridMon для Grid-специфічного моніторингу.'));
  
  body.push(h2('48.3. Job Submission'));
  body.push(p('Мова опису задач: JDL (Job Description Language) у gLite. RSL (Resource Specification Language) у Globus. Опис містить: вимоги до CPU/RAM, опис виконуваного файлу, вхідні/вихідні дані, обмеження часу. Брокер знаходить відповідний CE (Computing Element) та надсилає задачу.'));
  
  body.push(h2('48.4. Virtual Organization (VO)'));
  body.push(p('Базова соціальна одиниця Grid: набір осіб та установ, що розподіляють ресурси для спільної мети. VOMS (Virtual Organization Membership Service) управляє членством та ролями. Приклади VO: ATLAS, CMS (фізика частинок), biomed (біомедицина).'));
  
  
  // ═══════════════════════════════════════════════════════════════════
  // ПИТАННЯ 50
  // ═══════════════════════════════════════════════════════════════════
  body.push(h1('Питання 50. Організація і управління розподіленими грід-ресурсами'));
  
  body.push(h2('50.1. Планування задач у Grid'));
  body.push(p('Grid-планування (Grid Scheduling) — складніше, ніж у кластері, через: гетерогенність ресурсів, автономність організацій, динамічна доступність, географічна розподіленість. Дворівневе планування: локальний планувальник (SLURM, PBS) + глобальний Grid-брокер.'));
  body.push(p('Стратегії: First-Fit (перший придатний ресурс), Best-Fit (найкращий за параметрами), Min-Min (мінімізація makespan для малих задач), Max-Min (великі задачі першими).'));
  
  body.push(h2('50.2. Load Balancing у Grid'));
  body.push(p('Балансування навантаження між вузлами Grid-системи: статичне (розподіл при поданні), динамічне (міграція задач між вузлами). Проблема: задача не може легко мігрувати між вузлами різних організацій через безпеку та архітектурні відмінності. Work stealing у рамках VO — можливий підхід.'));
  
  body.push(h2('50.3. Data Management у Grid'));
  body.push(p('Управління даними в Grid включає:'));
  body.push(bullet('Реплікація: копії даних на кількох вузлах для надійності та швидкого доступу. LCG File Catalog (LFC) зберігає метадані та місцезнаходження реплік.'));
  body.push(bullet('GridFTP: протокол передачі файлів, оптимізований для Grid. Паралельні потоки, 3rd-party transfer (A → B через C), часткова передача.'));
  body.push(bullet('Data Staging: автоматичне переміщення вхідних даних на вузол перед запуском задачі.'));
  
  body.push(h2('50.4. Fault Tolerance'));
  body.push(p('Grid-системи мають низьку надійність окремих вузлів. Методи відмовостійкості:'));
  body.push(bullet('Checkpointing: збереження стану задачі та перезапуск при відмові.'));
  body.push(bullet('Replication: запуск задачі на кількох вузлах, прийняти перший результат (quorum).'));
  body.push(bullet('Task Migration: переміщення задачі з відмовленого вузла.'));
  body.push(bullet('Pilot Jobs (late binding): "пілотна задача" резервує ресурс, реальна задача запускається лише коли ресурс готовий, — уникає черги.'));
  
  body.push(h2('50.5. Сучасні тенденції: Grid + Cloud'));
  body.push(p('Традиційний Grid замінюється або доповнюється хмарними технологіями. EGI (European Grid Infrastructure) інтегрується з OpenStack. WLCG (LHC Computing Grid) використовує комбінацію Grid та Cloud ресурсів. Kubernetes для оркестрації контейнерів замінює GRAM. HTCondor + AWS/GCP — гібридний Grid-Cloud.'));
  
  body.push(h2('50.6. Управління безпекою'));
  body.push(p('Кожна організація в Grid має власну security policy. GSI забезпечує: mutual authentication (обидві сторони підтверджують ідентичність), delegation (proxy certificate дозволяє Grid-системі діяти від імені користувача), authorization (VOMS roles визначають права доступу до ресурсів).'));
  
  
  // ─── BUILD DOCUMENT ──────────────────────────────────────────────────────────
  
  // Title page
  const titlePage = [
    new Paragraph({ children: [new TextRun('')], spacing: { before: 3000, after: 0 } }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ text: 'ТЕХНОЛОГІЯ ПАРАЛЕЛЬНИХ ОБЧИСЛЕНЬ', bold: true, size: 40, font: 'Arial', color: '1F3864' })],
      spacing: { before: 0, after: 240 },
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ text: 'Відповіді на екзаменаційні питання', size: 32, font: 'Arial', color: '444444' })],
      spacing: { before: 0, after: 200 },
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ text: 'Парний варіант (питання 2, 4, 6, 8, ..., 50)', size: 26, font: 'Arial', color: '666666', italics: true })],
      spacing: { before: 0, after: 200 },
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ text: 'Мова реалізації: Java  |  MPI / OpenMPI', size: 24, font: 'Arial', color: '888888' })],
      spacing: { before: 0, after: 0 },
    }),
  ];
  
  const pageProperties = {
    page: {
      size: { width: A4_W, height: A4_H },
      margin: { top: MARGIN, right: MARGIN, bottom: MARGIN, left: 1800 },
    }
  };
  
  const doc = new Document({
    styles: {
      default: { document: { run: { font: 'Arial', size: 24 } } },
      paragraphStyles: [
        {
          id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true,
          run: { size: 32, bold: true, font: 'Arial', color: '1F3864' },
          paragraph: { spacing: { before: 320, after: 160 }, outlineLevel: 0 },
        },
        {
          id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true,
          run: { size: 28, bold: true, font: 'Arial', color: '2E75B6' },
          paragraph: { spacing: { before: 220, after: 100 }, outlineLevel: 1 },
        },
      ],
    },
    numbering: {
      config: [
        {
          reference: 'bullets',
          levels: [{
            level: 0, format: LevelFormat.BULLET, text: '-',
            alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 720, hanging: 360 } } },
          }],
        },
      ],
    },
    sections: [
      // Титульна сторінка
      {
        properties: pageProperties,
        children: titlePage,
      },
      // Основний контент
      {
        properties: pageProperties,
        headers: {
          default: new Header({
            children: [
              new Paragraph({
                children: [new TextRun({ text: 'Технологія паралельних обчислень — парні питання (2..50)', size: 18, font: 'Arial', color: '888888' })],
                border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: 'CCCCCC', space: 4 } },
                spacing: { after: 160 },
              }),
            ],
          }),
        },
        footers: {
          default: new Footer({
            children: [
              new Paragraph({
                alignment: AlignmentType.RIGHT,
                children: [new TextRun({ text: 'Java | MPI | Grid', size: 18, font: 'Arial', color: 'AAAAAA' })],
                border: { top: { style: BorderStyle.SINGLE, size: 4, color: 'CCCCCC', space: 4 } },
              }),
            ],
          }),
        },
        children: body,
      },
    ],
  });
  
  Packer.toBuffer(doc).then(buf => {
    fs.writeFileSync('./docs/ev4.docx', buf);
  }).catch(err => {
    console.error('Помилка:', err);
  });