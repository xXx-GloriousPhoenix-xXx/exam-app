// ============================================================
// Запуск: node generate_exam_even.js
// Потрібно: npm install docx
// Результат: exam_answers_even.docx
// ============================================================

const {
    Document, Packer, Paragraph, TextRun,
    HeadingLevel, AlignmentType, LevelFormat,
    PageBreak, BorderStyle
  } = require('docx');
  const fs = require('fs');
  
  // ============================================================
  // ДАНІ: питання та розгорнуті відповіді
  // ============================================================
  const questions = [
  
    // ----------------------------------------------------------
    {
      id: 2,
      title: "2. Класифікація паралельних обчислювальних систем",
      answer: [
        { heading: "Необхідність класифікації" },
        { text: "Паралельні обчислювальні системи охоплюють широкий спектр архітектур — від кількох ядер у смартфоні до суперкомп'ютерів з мільйонами процесорів. Класифікація дозволяє систематизувати ці системи, вибирати правильні інструменти та алгоритми для конкретного класу та прогнозувати продуктивність. Найбільш поширеними класифікаційними критеріями є: організація потоків інструкцій та даних (таксономія Фліна); архітектура пам'яті (спільна vs розподілена); топологія мережі взаємозв'язків; призначення системи." },
  
        { heading: "Таксономія Фліна (Flynn's Taxonomy)" },
        { text: "Майкл Флін у 1966 році запропонував класифікацію архітектур за двома вимірами: кількістю потоків інструкцій (Instruction Streams) та потоків даних (Data Streams). Це дало чотири класи: SISD, SIMD, MISD, MIMD." },
        { text: "SISD (Single Instruction, Single Data) — класична фон Нейманівська архітектура. Один процесор виконує одну послідовність інструкцій над одним потоком даних. Всі класичні алгоритми написані для SISD. Сьогодні зустрічається у вбудованих мікроконтролерах та дуже простих пристроях." },
        { text: "SIMD (Single Instruction, Multiple Data) — один потік інструкцій, але кожна інструкція одночасно оброблює масив даних. Класичні приклади: векторні розширення процесорів Intel/AMD (SSE — 128-бітні операції, AVX — 256-бітні, AVX-512 — 512-бітні); GPU (графічні процесори) — тисячі шейдерних ядер виконують одну програму паралельно над мільйонами пікселів; DSP-процесори для обробки сигналів. SIMD ідеальний для матричних операцій, фільтрації зображень, машинного навчання (операції над тензорами). Java Vector API (JEP 338/414/417, introduced in Java 16 as incubator) надає доступ до апаратних SIMD інструкцій." },
        { text: "MISD (Multiple Instruction, Single Data) — кілька процесорів виконують різні програми над тими ж даними. На практиці зустрічається рідко. Приклади: конвеєрні архітектури (кожна стадія конвеєра — окремий 'процесор'); fault-tolerant системи, де кілька різних алгоритмів обробляють одні й ті самі дані для перехресної перевірки (авіаційна та космічна електроніка)." },
        { text: "MIMD (Multiple Instruction, Multiple Data) — найпоширеніша категорія сучасних паралельних систем. Кілька процесорів незалежно виконують різні програми над різними даними. Багатоядерні CPU, кластери, суперкомп'ютери — все це MIMD. Java-програми з Thread, ForkJoin, ExecutorService — реалізація MIMD на рівні ПЗ." },
  
        { heading: "Класифікація за архітектурою пам'яті" },
        { text: "Системи зі спільною пам'яттю (Shared Memory, SM) — всі процесори мають доступ до єдиного адресного простору. Комунікація між потоками відбувається через читання/запис спільних змінних. Переваги: простота програмування; низька затримка доступу до даних. Недоліки: масштабованість обмежена (шина пам'яті стає вузьким місцем); необхідна синхронізація. Підкатегорії: UMA (Uniform Memory Access) — всі процесори мають однаковий час доступу до будь-якої ділянки пам'яті (SMP — Symmetric MultiProcessing). NUMA (Non-Uniform Memory Access) — час доступу залежить від того, яким процесором яка пам'ять була виділена. Кожен процесор має 'свою' локальну пам'ять та може звертатися до пам'яті інших через interconnect, але повільніше." },
        { text: "Системи з розподіленою пам'яттю (Distributed Memory, DM) — кожен вузол (процесор) має власну приватну пам'ять. Зв'язок між вузлами — виключно через явний обмін повідомленнями (message passing). Приклади: кластери серверів, суперкомп'ютери (Fugaku, Frontier). Переваги: практично необмежена масштабованість (тисячі/мільйони вузлів); немає конфлікту за шину пам'яті. Недоліки: складніше програмування (треба явно передавати дані); висока затримка міжвузлових комунікацій. Інструменти: MPI (Message Passing Interface), MPJ Express для Java." },
        { text: "Гібридні системи — поєднують спільну та розподілену пам'ять. Кожен вузол кластера є SMP/NUMA-машиною (спільна пам'ять всередині вузла), а вузли між собою мають розподілену пам'ять та спілкуються через мережу. Сучасні суперкомп'ютери є гібридними: MPI між вузлами + OpenMP/Java threads всередині вузла + CUDA/OpenCL на GPU." },
  
        { heading: "Класифікація за топологією мережі" },
        { text: "Топологія визначає, як процесори з'єднані між собою. Шина (Bus): проста, дешева, але стає вузьким місцем при збільшенні кількості процесорів. Повний граф (Fully Connected): кожен з кожним; O(N²) з'єднань — дорого. Гіперкуб (Hypercube): D-вимірний куб, 2^D вузлів, D з'єднань на вузол, діаметр D — хороший компроміс між вартістю та продуктивністю. Дерево (Tree/Fat-Tree): ієрархічна структура; Fat-Tree (потовщене дерево) є стандартом у HPC кластерах (InfiniBand Fat-Tree у більшості TOP500 систем). Торус (Torus): 3D або 6D тор — використовується в IBM Blue Gene, Cray суперкомп'ютерах. Crossbar switch: матриця комутаторів; O(N²) ціна, але O(1) затримка." },
  
        { heading: "Класифікація за масштабом" },
        { text: "Multicore (багатоядерні процесори): 2–128 ядер на одному кристалі; спільна кеш L3; UMA або NUMA в межах пакету. Manycore (масово-паралельні): GPU (тисячі ядер); Intel Xeon Phi (до 72 ядер). Кластер: від 10 до тисяч серверів; Fast Ethernet, InfiniBand; MPI. Суперкомп'ютер: 100,000+ ядер; спеціалізована мережа; сотні ПБ пам'яті. Grid / Cloud: мільйони машин; інтернет-з'єднання; гетерогенне обладнання." }
      ]
    },
  
    // ----------------------------------------------------------
    {
      id: 4,
      title: "4. Системи із загальною та розподіленою пам'яттю",
      answer: [
        { heading: "Системи зі спільною пам'яттю" },
        { text: "У системах зі спільною пам'яттю (Shared Memory Systems) всі процесори/ядра мають прямий доступ до єдиного фізичного адресного простору. Комунікація між потоками відбувається неявно — через читання та запис спільних об'єктів у пам'яті. Саме ця модель реалізована в Java: усі потоки одного процесу поділяють heap, статичні поля та відкриті ресурси." },
        { text: "SMP (Symmetric MultiProcessing) — найпростіша форма shared memory. Всі процесори рівноправні та мають однаковий час доступу до будь-якої ділянки RAM (UMA — Uniform Memory Access). Типовий приклад — звичайний багатоядерний ноутбук або сервер з одним процесорним сокетом. Апаратний контролер пам'яті та когерентний кеш-протокол (MESI або MOESI) забезпечують узгодженість кешів різних ядер." },
        { text: "NUMA (Non-Uniform Memory Access) — кожен процесорний вузол (socket) має локальну пам'ять, до якої має швидкий доступ. Доступ до пам'яті іншого сокету відбувається через interconnect (QPI у Intel, Infinity Fabric у AMD) і є повільнішим у 2–4 рази. Сучасні багатосокетні сервери є NUMA-машинами. Java-програми на NUMA можуть страждати від NUMA miss — потоки звертаються до 'чужої' пам'яті. Рішення: NUMA-aware memory allocation, прив'язка потоків до NUMA-вузлів (numactl у Linux)." },
  
        { heading: "Когерентність кешу — ключова проблема" },
        { text: "Кожне ядро має приватний кеш (L1, L2). Якщо два ядра кешують одну й ту саму cache line і одне з них змінює значення, кеш іншого стає застарілим. Протокол когерентності кешу (MESI: Modified, Exclusive, Shared, Invalid) автоматично підтримує узгодженість, але ціною продуктивності. False Sharing — ситуація, коли два потоки незалежно змінюють різні змінні, що знаходяться в одній cache line (64 байти). Кожна зміна інвалідує кеш іншого ядра, хоча логічного конфлікту немає. Це типова прихована причина деградації продуктивності паралельних Java-програм." },
        { code: `// ПРИКЛАД FALSE SHARING:
  class BadShared {
      long threadA_counter;  // використовується потоком A
      long threadB_counter;  // використовується потоком B
      // обидва поля в одній cache line — false sharing!
  }
  
  // ВИПРАВЛЕННЯ з @Contended (Java 8+):
  // JVM прапорець: -XX:-RestrictContended
  class GoodShared {
      @sun.misc.Contended volatile long threadA_counter;
      @sun.misc.Contended volatile long threadB_counter;
  }
  
  // Альтернативно — ручний padding (портативно):
  class PaddedCounter {
      long value;
      long p1, p2, p3, p4, p5, p6, p7; // 7 * 8 = 56 байт padding
  }` },
  
        { heading: "Системи з розподіленою пам'яттю" },
        { text: "У системах з розподіленою пам'яттю (Distributed Memory Systems) кожен вузол має власну приватну пам'ять. Жоден процес не може напряму читати або записувати пам'ять іншого вузла — обмін відбувається виключно через явну передачу повідомлень по мережі. Це фундаментально інша модель програмування порівняно зі shared memory." },
        { text: "Ключові характеристики: масштабованість — можна об'єднати тисячі вузлів без конфліктів за шину; відмовостійкість — відмова одного вузла не обов'язково призупиняє решту; складність програмування — треба явно розподіляти дані та планувати комунікації; затримка мережі — навіть InfiniBand HDR (200 Gbps) має затримку ~1–2 мікросекунди vs ~100 наносекунд для доступу до RAM." },
  
        { heading: "MPI як стандарт для distributed memory" },
        { text: "MPI (Message Passing Interface) є де-факто стандартом для програмування систем з розподіленою пам'яттю. Основні операції: MPI_Send/MPI_Recv (point-to-point); MPI_Bcast (один надсилає всім); MPI_Scatter (root розподіляє дані між усіма); MPI_Gather (всі надсилають root'у); MPI_Reduce/MPI_AllReduce (колективна операція з редукцією). У Java використовується MPJ Express або Open MPI з Java bindings." },
        { code: `// СХЕМА розподілу даних у distributed memory (MPJ Express):
  // Кожен MPI-процес бачить тільки свою частину матриці A
  
  MPI.Init(args);
  int rank = MPI.COMM_WORLD.Rank();
  int size = MPI.COMM_WORLD.Size();
  int N = 1000;
  int rowsPerProc = N / size;
  
  // Власна пам'ять кожного процесу — тільки його рядки:
  double[] localA = new double[rowsPerProc * N];
  double[] B = new double[N * N]; // копія B у кожному процесі
  
  // Root розсилає частини A:
  MPI.COMM_WORLD.Scatter(
      (rank == 0) ? A_flat : null, 0, rowsPerProc * N, MPI.DOUBLE,
      localA, 0, rowsPerProc * N, MPI.DOUBLE, 0
  );
  // Root broadcast B всім:
  MPI.COMM_WORLD.Bcast(B, 0, N * N, MPI.DOUBLE, 0);
  
  // Кожен обчислює свої рядки — ЛОКАЛЬНО, без комунікацій:
  double[] localC = multiplyRows(localA, B, rowsPerProc, N);
  
  // Збір результатів:
  MPI.COMM_WORLD.Gather(
      localC, 0, rowsPerProc * N, MPI.DOUBLE,
      (rank == 0) ? C_flat : null, 0, rowsPerProc * N, MPI.DOUBLE, 0
  );
  MPI.Finalize();` },
  
        { heading: "Порівняння shared vs distributed memory" },
        { text: "Shared memory: простота написання (просто спільні змінні); низька затримка (~100 нс); обмежена масштабованість (практично до ~256 ядер); потрібна синхронізація (synchronized, volatile, locks). Distributed memory: складніше програмування (явна передача даних); вища затримка (~1–1000 мкс); необмежена масштабованість (мільйони процесорів); природна ізоляція (без race conditions). Сучасний підхід — гібридний: MPI між вузлами кластера (distributed memory) + Java threads/ForkJoin всередині вузла (shared memory)." }
      ]
    },
  
    // ----------------------------------------------------------
    {
      id: 6,
      title: "6. Багатопоточна технологія Java",
      answer: [
        { heading: "Java як платформа для конкурентного програмування" },
        { text: "Java від самого початку (JDK 1.0, 1995) була розроблена з підтримкою багатопоточності як першокласного механізму. На відміну від C/C++, де потоки були бібліотечним розширенням (POSIX threads, Win32 threads), Java вбудувала підтримку паралелізму у мову та стандартну бібліотеку. Ключові елементи: клас Thread та інтерфейс Runnable (з JDK 1.0); ключові слова synchronized та volatile (JDK 1.0); Java Memory Model — специфікація видимості змін між потоками (формалізована у JDK 5 через JSR-133); пакет java.util.concurrent з ExecutorService, ConcurrentCollections, Lock API, AtomicVariables (JDK 5, JSR-166); ForkJoinPool (JDK 7); CompletableFuture (JDK 8); Virtual Threads / Project Loom (JDK 21)." },
  
        { heading: "Три способи створення задачі для паралельного виконання" },
        { text: "Спосіб 1 — успадкування від Thread. Клас MyTask extends Thread, перевизначаємо run(). Недолік: Java не підтримує множинне успадкування — клас не може одночасно успадковувати Thread і ще щось. Спосіб 2 — реалізація Runnable (рекомендований). MyTask implements Runnable, передаємо в new Thread(runnable). Відокремлює задачу від механізму виконання. Спосіб 3 — реалізація Callable<V>. Як Runnable, але може повертати результат типу V та кидати checked exceptions. Використовується з ExecutorService.submit() та Future<V>." },
        { code: `// Спосіб 2: Runnable + Thread
  Runnable task = () -> {
      System.out.println("Виконується у: " +
          Thread.currentThread().getName());
  };
  Thread t = new Thread(task, "my-worker");
  t.setDaemon(false);
  t.setPriority(Thread.NORM_PRIORITY);
  t.start();
  t.join(); // чекаємо завершення
  
  // Спосіб 3: Callable + Future
  ExecutorService exec = Executors.newFixedThreadPool(4);
  Future<Double> future = exec.submit(() -> Math.sqrt(2.0));
  Double result = future.get(5, TimeUnit.SECONDS);
  exec.shutdown();` },
  
        { heading: "Java Memory Model (JMM)" },
        { text: "JMM визначає правила видимості змін між потоками. Без JMM гарантій JIT-компілятор та CPU можуть переупорядковувати інструкції, зберігати значення в регістрах або кеші без запису в пам'ять. Happens-before відношення гарантує, що зміни одного потоку видимі іншому: Thread.start() hb першої дії потоку; unlock() hb наступного lock() того ж монітора; запис volatile hb читання того ж volatile; Thread.join() hb після join(). volatile гарантує видимість та заборону reordering, але НЕ атомарність складених операцій. synchronized гарантує і видимість, і взаємне виключення." },
  
        { heading: "Пакет java.util.concurrent" },
        { text: "java.util.concurrent (JUC) — потужний тулкіт для конкурентного програмування, що значно перевершує 'сирий' Thread API. Executor Framework: Executor, ExecutorService, ScheduledExecutorService — абстракція пулу потоків. Конкурентні колекції: ConcurrentHashMap (thread-safe HashMap без глобального блокування); CopyOnWriteArrayList (безпечний для читання списку); BlockingQueue (LinkedBlockingQueue, ArrayBlockingQueue, PriorityBlockingQueue, SynchronousQueue) — ключова структура для Producer-Consumer паттерну. Синхронізатори: CountDownLatch (одноразовий бар'єр); CyclicBarrier (перезапускається бар'єр); Semaphore (лічильник дозволів); Phaser (гнучкий, замінює CountDownLatch+CyclicBarrier); Exchanger (двосторонній обмін між двома потоками). Atomic variables: AtomicInteger, AtomicLong, AtomicReference, LongAdder — lock-free операції через CAS." },
  
        { heading: "Virtual Threads (Project Loom, Java 21)" },
        { text: "Virtual Threads — революційне нововведення Java 21. Звичайні (platform) потоки — це обгортки над ОС-потоками, яких можна створити ~10,000. Virtual threads — надлегкі потоки, яких можна створити мільйони. JVM самостійно мультиплексує virtual threads на малу кількість platform threads (carrier threads). При блокуванні (I/O, sleep) virtual thread знімається з carrier thread без блокування ОС-потоку. Це дозволяє писати синхронний за виглядом код, що масштабується як асинхронний." },
        { code: `// Virtual Thread (Java 21):
  try (var executor = Executors.newVirtualThreadPerTaskExecutor()) {
      for (int i = 0; i < 1_000_000; i++) {
          executor.submit(() -> {
              Thread.sleep(Duration.ofSeconds(1)); // не блокує carrier thread!
              return processRequest();
          });
      }
  } // auto shutdown` },
  
        { heading: "Паттерн Producer-Consumer у Java" },
        { text: "Producer-Consumer — один з найважливіших паттернів конкурентного програмування. Producers генерують задачі/дані, Consumers їх обробляють. BlockingQueue є природним буфером між ними: put() блокується при заповненій черзі; take() блокується при порожній черзі. Це забезпечує автоматичне регулювання швидкості (backpressure)." },
        { code: `BlockingQueue<Task> queue = new LinkedBlockingQueue<>(100);
  
  // Producer:
  Thread producer = new Thread(() -> {
      try {
          while (!Thread.currentThread().isInterrupted()) {
              queue.put(generateTask()); // блокується якщо черга повна
          }
      } catch (InterruptedException e) {
          Thread.currentThread().interrupt();
      }
  });
  
  // Consumer:
  Thread consumer = new Thread(() -> {
      try {
          while (!Thread.currentThread().isInterrupted()) {
              Task t = queue.take(); // блокується якщо черга порожня
              process(t);
          }
      } catch (InterruptedException e) {
          Thread.currentThread().interrupt();
      }
  });` }
      ]
    },
  
    // ----------------------------------------------------------
    {
      id: 8,
      title: "8. Стани потоку та переходи потоку з одного стану в інший",
      answer: [
        { heading: "Перелік станів потоку (Thread.State)" },
        { text: "Enum Thread.State у Java визначає 6 можливих станів: NEW, RUNNABLE, BLOCKED, WAITING, TIMED_WAITING, TERMINATED. Поточний стан отримується через thread.getState(). Важливо розуміти, що стани Java — це логічні стани JVM; реальний стан ОС-потоку може відрізнятися (наприклад, RUNNABLE у Java може відповідати як running, так і ready у планувальнику ОС)." },
  
        { heading: "NEW — потік створено, але не запущено" },
        { text: "Потік перебуває у стані NEW від моменту виклику конструктора Thread() до виклику start(). У цьому стані JVM ще не виділила системні ресурси для потоку, ОС-потік ще не існує. Переходи з NEW: тільки у RUNNABLE після виклику start(). Не можна викликати start() двічі — кинеться IllegalThreadStateException." },
        { code: `Thread t = new Thread(() -> System.out.println("Hello"));
  System.out.println(t.getState()); // NEW
  t.start();
  System.out.println(t.getState()); // RUNNABLE (або вже TERMINATED)` },
  
        { heading: "RUNNABLE — виконується або готовий до виконання" },
        { text: "Стан RUNNABLE об'єднує два фізичних стани ОС: running (потік зараз виконується на ядрі процесора) та ready (потік готовий, але очікує виділення процесорного часу планувальником). Java не розрізняє ці два стани — обидва є RUNNABLE. Потік входить у RUNNABLE після start(), після повернення з sleep()/wait()/join(), після отримання монітора (з BLOCKED) або після виклику LockSupport.unpark(). Переходи з RUNNABLE: у BLOCKED (намагається увійти в synchronized, але монітор зайнятий); у WAITING (wait(), join() без таймауту); у TIMED_WAITING (sleep(n), wait(n), join(n)); у TERMINATED (метод run() завершився)." },
  
        { heading: "BLOCKED — очікування монітора" },
        { text: "Потік перебуває у BLOCKED, коли намагається увійти у synchronized блок або метод, але монітор вже захоплений іншим потоком. Або коли потік прокидається після wait() і намагається знову захопити монітор. Вихід з BLOCKED — тільки коли монітор стає доступним і планувальник надає потоку монітор. Стан BLOCKED є одним з джерел deadlock: якщо два потоки взаємно очікують монітори один одного, обидва застрягають у BLOCKED назавжди." },
  
        { heading: "WAITING — безстрокове очікування" },
        { text: "Потік переходить у WAITING при виклику: Object.wait() (без таймауту) — всередині synchronized блоку; Thread.join() (без таймауту) — очікування завершення іншого потоку; LockSupport.park() — примітивна операція паркування (використовується всередині AbstractQueuedSynchronizer). Вихід з WAITING відбувається при: Object.notify() або Object.notifyAll() (з WAITING через wait() — переходить до BLOCKED, щоб знову захопити монітор); Thread.interrupt() — кидає InterruptedException; LockSupport.unpark(thread) — пробуджує запаркований потік. КРИТИЧНО: wait() завжди використовувати в циклі while (!condition) { wait(); } через spurious wakeup." },
  
        { heading: "TIMED_WAITING — очікування з таймаутом" },
        { text: "Аналог WAITING, але з часовим обмеженням. Настає при: Thread.sleep(millis) або Thread.sleep(Duration); Object.wait(millis); Thread.join(millis); LockSupport.parkNanos(nanos) / parkUntil(deadline). Вихід: після закінчення таймауту; або ті ж умови, що для WAITING (notify, interrupt, unpark)." },
  
        { heading: "TERMINATED — потік завершений" },
        { text: "Потік переходить у TERMINATED коли метод run() повертає значення (нормальне завершення) або коли з run() вилітає необроблений виняток. Після TERMINATED потік не може бути перезапущений — будь-яка спроба викликати start() знову призведе до IllegalThreadStateException. thread.isAlive() повертає false." },
  
        { heading: "Діаграма переходів між станами" },
        { text: "NEW --[start()]--> RUNNABLE <--> BLOCKED (захоплення монітора). RUNNABLE --[wait()]--> WAITING --[notify()]--> BLOCKED --[монітор вільний]--> RUNNABLE. RUNNABLE --[sleep(n)]--> TIMED_WAITING --[час закінчився]--> RUNNABLE. RUNNABLE --[run() завершено]--> TERMINATED." },
  
        { heading: "Практична діагностика стану потоків" },
        { code: `// Отримання стану конкретного потоку:
  Thread worker = new Thread(task);
  Thread.State state = worker.getState();
  
  // Дамп всіх потоків (для діагностики deadlock):
  Map<Thread, StackTraceElement[]> allTraces = Thread.getAllStackTraces();
  for (Map.Entry<Thread, StackTraceElement[]> entry : allTraces.entrySet()) {
      Thread t = entry.getKey();
      System.out.println(t.getName() + " [" + t.getState() + "]");
      for (StackTraceElement ste : entry.getValue()) {
          System.out.println("    " + ste);
      }
  }
  
  // jstack <pid> — утиліта командного рядка для аналізу deadlock
  // VisualVM, JConsole — GUI для моніторингу потоків` }
      ]
    },
  
    // ----------------------------------------------------------
    {
      id: 10,
      title: "10. Алгоритми паралельного множення матриць",
      answer: [
        { heading: "Постановка задачі та важливість" },
        { text: "Множення матриць C = A × B, де A розміром M×K, B розміром K×N, C розміром M×N. Елемент C[i][j] = Σ(p=0..K-1) A[i][p] * B[p][j]. Послідовна складність O(M×K×N), або O(N³) для квадратних матриць. При N=2000 це 8×10⁹ операцій. Множення матриць є основою: нейронних мереж (forward/backward pass — матричне множення); наукових симуляцій; комп'ютерної графіки; розв'язання систем лінійних рівнянь. Саме тому ефективні паралельні алгоритми множення матриць є центральним завданням курсу." },
  
        { heading: "Стрічковий алгоритм (Row-Stripe Decomposition)" },
        { text: "Найпростіший і найпоширеніший підхід. Матриця A ділиться на горизонтальні смуги (стрічки) по рядках. Кожен потік/процес отримує свою стрічку рядків A та повну матрицю B, обчислює відповідні рядки C. Потоки повністю незалежні — немає синхронізацій під час обчислення (embarrassingly parallel). Синхронізація потрібна лише після завершення (join для потоків, Gather для MPI)." },
        { code: `// Стрічковий алгоритм з Thread (Java):
  public class StripedMatMul {
      public static double[][] multiply(double[][] A, double[][] B, int p)
              throws InterruptedException {
          int M = A.length, K = B.length, N = B[0].length;
          double[][] C = new double[M][N];
          Thread[] threads = new Thread[p];
          int strip = (M + p - 1) / p;
  
          for (int t = 0; t < p; t++) {
              final int start = t * strip;
              final int end = Math.min(start + strip, M);
              threads[t] = new Thread(() -> {
                  for (int i = start; i < end; i++)
                      for (int j = 0; j < N; j++)
                          for (int k = 0; k < K; k++)
                              C[i][j] += A[i][k] * B[k][j];
              });
              threads[t].start();
          }
          for (Thread th : threads) th.join();
          return C;
      }
  }` },
  
        { heading: "Блоковий алгоритм (Block / Tile Decomposition)" },
        { text: "Матриці A, B і C ділять на квадратні або прямокутні блоки розміром b×b. Кожен потік/процес обробляє один або кілька блоків C. Перевага: cache-friendly — блоки вміщуються у кеш L2/L3, що значно скорочує кількість cache miss. Алгоритм природньо обробляє блок-стовпці B та блок-рядки A. При оптимальному виборі b (зазвичай 32–128 для сучасних CPU) блоковий алгоритм може бути у 5–10 разів швидшим за наївний." },
  
        { heading: "Алгоритм Фокса (Fox's Algorithm / BSP)" },
        { text: "Алгоритм Фокса — метод паралельного множення матриць для топологічної сітки процесів √p × √p (distributed memory). На кожному кроці i (від 0 до √p-1): diagonal_block = A[row][( row+i ) mod √p]; broadcast diagonal_block уздовж рядка сітки; кожен процес множить отриманий блок A на свій локальний блок B і додає до часткового C; B зсувається вгору по стовпцю на 1 позицію. Після √p кроків кожен процес містить фінальний блок C. Перевага: обсяг переданих даних O(N²/√p) — оптимальний для 2D-сіток." },
  
        { heading: "Алгоритм Кеннона (Cannon's Algorithm)" },
        { text: "Алгоритм Кеннона (1969) розроблений для √p × √p сітки процесів. Спочатку виконується початковий зсув: рядок i матриці A зсувається вліво на i позицій; стовпець j матриці B зсувається вгору на j позицій. Потім √p ітерацій: local_C += local_A * local_B; зсунути всі блоки A вліво на 1; зсунути всі блоки B вгору на 1. На відміну від Фокса, Кеннон не потребує broadcast — лише point-to-point комунікації (Sendrecv), що ефективніше для великих сіток." },
  
        { heading: "ForkJoin реалізація (рекурсивне розбиття)" },
        { code: `public class ForkJoinMatMul extends RecursiveAction {
      static final int THRESHOLD = 64;
      double[][] A, B, C;
      int rowStart, rowEnd, N;
  
      protected void compute() {
          if (rowEnd - rowStart <= THRESHOLD) {
              for (int i = rowStart; i < rowEnd; i++)
                  for (int j = 0; j < N; j++)
                      for (int k = 0; k < N; k++)
                          C[i][j] += A[i][k] * B[k][j];
              return;
          }
          int mid = (rowStart + rowEnd) / 2;
          invokeAll(
              new ForkJoinMatMul(A, B, C, rowStart, mid, N),
              new ForkJoinMatMul(A, B, C, mid, rowEnd, N)
          );
      }
  }` },
  
        { heading: "Spring Boot застосунок для множення матриць" },
        { text: "Spring Boot дозволяє побудувати REST-сервіс для обчислення добутку матриць. Клієнт надсилає POST /multiply з тілом {matrixA: [[...]], matrixB: [[...]], threads: 8}. Сервер виконує паралельне обчислення та повертає resultMatrix. Можливості: асинхронне виконання (@Async + CompletableFuture); відстеження прогресу через Server-Sent Events або WebSocket; кешування результатів (Spring Cache + Redis); балансування навантаження між кількома інстанціями." },
        { code: `@RestController
  @RequestMapping("/api/matrix")
  public class MatrixController {
  
      @PostMapping("/multiply")
      public ResponseEntity<double[][]> multiply(
              @RequestBody MatrixRequest request) {
          double[][] result = matrixService.multiplyParallel(
              request.getA(), request.getB(),
              Runtime.getRuntime().availableProcessors()
          );
          return ResponseEntity.ok(result);
      }
  }` }
      ]
    },
  
    // ----------------------------------------------------------
    {
      id: 12,
      title: "12. Алгоритм Фокса паралельного множення матриць",
      answer: [
        { heading: "Передумови та мотивація" },
        { text: "При розподіленому множенні матриць на кластері p процесів кожен процес повинен отримати необхідні дані, обчислити свою частину результату та повернути її. Наївний підхід (кожен процес отримує повні матриці) не масштабується: обсяг комунікацій O(N²) на кожний процес. Алгоритм Фокса (також відомий як BSP або Column-Broadcast алгоритм) мінімізує об'єм переданих даних шляхом блокової декомпозиції та поетапного broadcast." },
  
        { heading: "Підготовка: блокова декомпозиція" },
        { text: "Припускаємо: кількість процесів p = q×q (q — ціле). Процеси організовані у двовимірну сітку q×q. Матриці A, B, C (розміром N×N) ділять на q×q блоків розміром b×b, де b = N/q. Кожен процес (i, j) містить блоки A[i][j], B[i][j] та обчислює блок C[i][j]." },
  
        { heading: "Кроки алгоритму Фокса" },
        { text: "Алгоритм виконується у q ітерацій (кроки k = 0, 1, ..., q-1). На кожному кроці k: 1) Визначення активного блоку A: для рядка i сітки активний блок — A[i][(i+k) mod q]. 2) Broadcast у рядку: процес (i, (i+k) mod q) розсилає A[i][(i+k) mod q] усім процесам рядка i. 3) Обчислення: кожен процес (i,j) множить отриманий блок A_active на свій локальний блок B[i][j] та додає до C[i][j]. 4) Циклічний зсув B: всі блоки B у кожному стовпці j зсуваються вгору на 1 (процес (i,j) надсилає свій B процесу ((i-1+q) mod q, j))." },
  
        { heading: "Реалізація в Java (MPJ Express, концептуально)" },
        { code: `// Алгоритм Фокса для q*q процесів:
  // Кожен процес має свій rank та знає свої координати (row, col) у сітці
  
  public static void foxAlgorithm(
          double[] localA, double[] localB, double[] localC,
          int blockSize, int q, MPI.Comm rowComm, MPI.Comm colComm,
          int row, int col) throws MPIException {
  
      double[] tempA = new double[blockSize * blockSize];
  
      for (int k = 0; k < q; k++) {
          // Визначаємо root для broadcast у поточному рядку:
          int broadcastRoot = (row + k) % q;
  
          // Копіюємо або отримуємо блок A для broadcast:
          if (col == broadcastRoot) {
              System.arraycopy(localA, 0, tempA, 0, tempA.length);
          }
  
          // Broadcast активного блоку A у межах рядка:
          rowComm.Bcast(tempA, 0, tempA.length, MPI.DOUBLE, broadcastRoot);
  
          // Множення: localC += tempA * localB
          multiplyBlocks(tempA, localB, localC, blockSize);
  
          // Циклічний зсув localB вгору на 1 у стовпці:
          double[] recvB = new double[blockSize * blockSize];
          int sendTo   = (row - 1 + q) % q;
          int recvFrom = (row + 1) % q;
          colComm.Sendrecv(
              localB, 0, localB.length, MPI.DOUBLE, sendTo, 0,
              recvB,  0, recvB.length, MPI.DOUBLE, recvFrom, 0
          );
          System.arraycopy(recvB, 0, localB, 0, localB.length);
      }
  }
  
  // Допоміжна функція: множення двох квадратних блоків:
  static void multiplyBlocks(double[] A, double[] B, double[] C, int b) {
      for (int i = 0; i < b; i++)
          for (int j = 0; j < b; j++)
              for (int k = 0; k < b; k++)
                  C[i * b + j] += A[i * b + k] * B[k * b + j];
  }` },
  
        { heading: "Складність та аналіз" },
        { text: "Обчислювальна вартість: кожен процес виконує q множень блоків b×b: q × 2b³ = q × 2(N/q)³ = 2N³/q² = 2N³/p операцій. Комунікаційна вартість: broadcast у рядку — q разів по b² = q × N²/q² = N²/q елементів; зсув B — q разів по b² = N²/q елементів. Загальна комунікаційна вартість: O(N²/√p). Це оптимально для 2D сітки: кожен процес передає лише 2/√p частину своїх даних на кожному кроці. Ефективність: при великих N та правильно вибраному p — близька до ідеальної." },
  
        { heading: "Порівняння Фокс vs Кеннон" },
        { text: "Алгоритм Фокса: використовує broadcast у рядку (один-до-всіх у підгрупі); простіший у реалізації з MPI Comm_split та MPI_Bcast; початковий зсув не потрібен. Алгоритм Кеннона: лише point-to-point Sendrecv між сусідами; потребує початкового зсуву A і B; ефективніший при дуже великих сітках, де broadcast дорогий. На практиці при рівних умовах продуктивність порівнянна; вибір залежить від архітектури мережі кластера." }
      ]
    },
  
    // ----------------------------------------------------------
    {
      id: 14,
      title: "14. Призупинка та відновлення роботи потоку. Переривання дії потоку",
      answer: [
        { heading: "Застарілі методи: suspend(), resume(), stop()" },
        { text: "У ранніх версіях Java (1.0) існували методи Thread.suspend(), Thread.resume() та Thread.stop(). Всі вони оголошені застарілими (deprecated since JDK 1.1/1.2) та небезпечними. Thread.stop() знищував потік, кидаючи ThreadDeath у довільний момент — монітори звільнялися, але об'єкти залишалися в неконсистентному стані, що призводило до пошкодження даних. Thread.suspend() зупиняла потік, не звільняючи його монітори — потік, що очікує монітор, захоплений призупиненим потоком, застрягав назавжди (deadlock). Сучасна Java пропонує безпечні альтернативи." },
  
        { heading: "Thread.sleep() — пауза з таймаутом" },
        { text: "Thread.sleep(long millis) та Thread.sleep(long millis, int nanos) — статичні методи для призупинки поточного потоку на вказаний час. Важливо: sleep() НЕ звільняє захоплені монітори (synchronized). Потік переходить у стан TIMED_WAITING. Якщо потік перервано (interrupt()) під час sleep(), кидається InterruptedException." },
        { code: `try {
      System.out.println("Починаю паузу");
      Thread.sleep(2000); // 2 секунди
      System.out.println("Пауза завершена");
  } catch (InterruptedException e) {
      // Відновлюємо прапорець переривання!
      Thread.currentThread().interrupt();
      System.out.println("Потік перервано під час сну");
  }
  
  // Java 11+: Thread.sleep(Duration.ofSeconds(2));` },
  
        { heading: "Object.wait() та Object.notify() — умовне очікування" },
        { text: "wait() ПОВИНЕН викликатись у synchronized блоці. На відміну від sleep(), wait() звільняє монітор об'єкта на час очікування, дозволяючи іншим потокам захопити його. Потік переходить у WAITING або TIMED_WAITING. notify() будить один довільний потік, що очікує на тому ж моніторі; notifyAll() будить усі. Після пробудження потік повертається у BLOCKED (намагається знову захопити монітор) і тільки після отримання монітора — у RUNNABLE." },
        { code: `// ПРАВИЛЬНИЙ паттерн: завжди while, не if!
  synchronized (sharedObject) {
      while (!conditionIsTrue()) {
          sharedObject.wait(); // звільняє монітор, чекає notify
      }
      // умова виконана, можна продовжувати
      doWork();
  }
  
  // Сигналізація з іншого потоку:
  synchronized (sharedObject) {
      conditionIsTrue = true;
      sharedObject.notifyAll(); // будимо всіх очікуючих
  }` },
  
        { heading: "Thread.join() — очікування завершення потоку" },
        { text: "join() блокує поточний потік до завершення потоку, на якому викликано join(). join(long millis) — очікує не довше вказаного часу. Типове використання: запуск N worker потоків, потім join() для кожного, щоб дочекатися результатів перед продовженням." },
        { code: `Thread[] workers = new Thread[N];
  for (int i = 0; i < N; i++) { workers[i] = new Thread(task); workers[i].start(); }
  for (Thread w : workers) {
      try { w.join(); }
      catch (InterruptedException e) { Thread.currentThread().interrupt(); break; }
  }
  // Тут усі workers завершились` },
  
        { heading: "Механізм переривання: interrupt()" },
        { text: "interrupt() — єдиний рекомендований спосіб 'зупинки' потоку в Java. Це кооперативний механізм: interrupt() лише встановлює булевий прапорець переривання у потоці. Якщо потік зараз заблокований у sleep(), wait(), join() або іншому InterruptedException-методі — негайно кидається InterruptedException і прапорець скидається. Якщо потік не заблокований — він продовжує виконання, але прапорець залишається встановленим до явної перевірки." },
        { code: `Thread worker = new Thread(() -> {
      while (!Thread.currentThread().isInterrupted()) {
          try {
              Task t = queue.take(); // BLOCKING — кине InterruptedException
              process(t);
          } catch (InterruptedException e) {
              Thread.currentThread().interrupt(); // відновлюємо прапорець
              break; // виходимо з циклу коректно
          }
      }
      cleanup(); // звільняємо ресурси
      System.out.println("Потік завершено коректно");
  });
  
  worker.start();
  Thread.sleep(5000);
  worker.interrupt(); // м'яко просимо зупинитися
  worker.join();      // чекаємо коректного завершення` },
  
        { heading: "LockSupport.park() та unpark() — низькорівневе паркування" },
        { text: "LockSupport є фундаментом реалізації всіх lock та synchronizer у JUC (AbstractQueuedSynchronizer). park() призупиняє поточний потік без захоплення монітора. unpark(thread) дозволяє конкретному потоку продовжити роботу. На відміну від wait/notify, unpark може бути викликано ДО park() — потік не призупиниться при наступному park(). Не рекомендується для прямого використання у прикладному коді — тільки для реалізації власних примітивів синхронізації." }
      ]
    },
  
    // ----------------------------------------------------------
    {
      id: 16,
      title: "16. Синхронізація в паралельних обчисленнях",
      answer: [
        { heading: "Навіщо потрібна синхронізація" },
        { text: "Синхронізація — координація між паралельними потоками/процесами для: взаємного виключення (mutual exclusion) — гарантії, що лише один потік одночасно виконує критичну секцію; узгодженості пам'яті (memory consistency) — гарантії видимості змін, зроблених одним потоком, іншими; координації (coordination) — синхронізації точок виконання між потоками (бар'єри, лічильники). Без синхронізації виникають race conditions — недетерміновані помилки, що залежать від планування потоків і надзвичайно важко відтворювані та відлагоджувані." },
  
        { heading: "Рівні синхронізації" },
        { text: "Апаратний рівень: атомарні інструкції процесора (CMPXCHG, XADD, LOCK-prefix у x86); memory barriers (MFENCE, SFENCE, LFENCE) для заборони переупорядкування. Рівень ОС: mutex (mutual exclusion object); semaphore; condition variable. Рівень JVM/Java: ключове слово synchronized (обгортає монітор); volatile (memory barrier без взаємного виключення); java.util.concurrent.locks (ReentrantLock, ReadWriteLock); java.util.concurrent.atomic (CAS-based операції). Рівень розподіленої системи: бар'єри MPI (MPI_Barrier); семафори та мютекси у рамках єдиного процесу на кластерному вузлі." },
  
        { heading: "Мютекс (Mutex) і synchronized" },
        { code: `// synchronized метод — монітор = this:
  public synchronized void criticalMethod() {
      // лише один потік може виконувати це одночасно
      sharedData++;
  }
  
  // synchronized блок — мінімальна критична секція:
  public void method() {
      doNonCriticalWork(); // паралельно
      synchronized (lockObject) {
          sharedData++; // критична секція
      }
      doMoreNonCriticalWork(); // паралельно
  }
  
  // Статичний метод — монітор = Class.class:
  public static synchronized void classLevelCritical() { }` },
  
        { heading: "Семафор (Semaphore)" },
        { text: "Семафор — лічильниковий примітив синхронізації. Бінарний семафор (0 або 1) аналогічний мютексу. Лічильниковий семафор використовується для обмеження кількості одночасних доступів до ресурсу: acquire() зменшує лічильник, блокуючись при 0; release() збільшує лічильник та прокидає один очікуючий потік." },
        { code: `// Обмеження кількості одночасних підключень до БД:
  Semaphore connectionPool = new Semaphore(10, true); // fair=true
  try {
      connectionPool.acquire();
      try {
          executeQuery();
      } finally {
          connectionPool.release();
      }
  } catch (InterruptedException e) {
      Thread.currentThread().interrupt();
  }` },
  
        { heading: "Бар'єр (Barrier) — точка синхронізації" },
        { text: "Бар'єр — точка, до якої всі паралельні виконавці (потоки або MPI-процеси) мають дійти, перш ніж будь-який з них рухається далі. Типове використання: ітеративні алгоритми, де кожна ітерація потребує результатів ВСІХ потоків попередньої ітерації. У Java: CyclicBarrier; CountDownLatch (одноразовий); Phaser (гнучкий). У MPI: MPI_Barrier()." },
        { code: `// CyclicBarrier: паралельна ітерація стрічкового алгоритму
  CyclicBarrier barrier = new CyclicBarrier(numThreads, () -> {
      // дія після того, як всі досягли бар'єру (одним потоком):
      exchangeBoundaryData();
  });
  
  // Кожен потік:
  for (int iteration = 0; iteration < maxIter; iteration++) {
      computeMyStripe(stripe, iteration);
      barrier.await(); // всі чекають тут
      // тепер граничні дані оновлено, можна продовжувати
  }` },
  
        { heading: "Синхронізація в MPI" },
        { text: "У MPI синхронізація відбувається через: MPI_Barrier() — явна точка синхронізації всіх процесів; колективні операції (Bcast, Scatter, Gather, Reduce) — неявно синхронізують беручі участь процеси; блокуючі Send/Recv — sender та receiver синхронізуються в точці обміну. Неблокуючі операції (MPI_Isend, MPI_Irecv + MPI_Wait) дозволяють накладати обчислення та комунікації." },
  
        { heading: "Lock-free синхронізація через CAS" },
        { text: "Compare-And-Swap (CAS) — атомарна апаратна операція: якщо поточне значення == expected, встановити newValue і повернути true, інакше нічого не змінювати і повернути false. CAS є основою AtomicInteger, AtomicReference тощо. Lock-free алгоритми гарантують прогрес системи (хоча б один потік завжди просувається), але не голодування конкретного потоку." },
        { code: `AtomicInteger counter = new AtomicInteger(0);
  counter.incrementAndGet();          // атомарний ++
  counter.compareAndSet(5, 10);       // якщо == 5, встановити 10
  
  // Lock-free linked list push (оптимістичний підхід):
  public void push(Node newNode) {
      Node head;
      do {
          head = headRef.get();
          newNode.next = head;
      } while (!headRef.compareAndSet(head, newNode));
      // retry якщо хтось інший змінив head між читанням і записом
  }` }
      ]
    },
  
    // ----------------------------------------------------------
    {
      id: 18,
      title: "18. Блокування об'єкту",
      answer: [
        { heading: "Монітор об'єкта в Java" },
        { text: "Кожен Java-об'єкт неявно має прихований монітор (intrinsic lock або monitor lock). Монітор — бінарний семафор, вбудований у заголовок об'єкта (object header). При вході у synchronized блок JVM виконує MONITORENTER байткод: якщо монітор вільний — захоплює його; якщо зайнятий — блокує поточний потік і ставить його в чергу очікування монітора. При виході — MONITOREXIT: звільняє монітор, прокидаючи один з очікуючих потоків (BLOCKED → конкуренція за монітор)." },
  
        { heading: "Реєнтерабельність (Reentrance)" },
        { text: "Java монітори є реєнтерабельними: якщо потік вже утримує монітор об'єкта і намагається знову захопити той самий монітор (наприклад, synchronized метод викликає інший synchronized метод того ж об'єкта) — він отримує його без блокування. JVM веде лічильник вкладеності: кожен MONITORENTER збільшує, кожен MONITOREXIT зменшує. Монітор звільняється тільки коли лічильник досягає 0." },
        { code: `public class ReentrantExample {
      public synchronized void outer() {
          System.out.println("outer: захопив монітор (count=1)");
          inner(); // виклик іншого synchronized методу того ж об'єкта
      }
  
      public synchronized void inner() {
          // Той самий потік знову захоплює той самий монітор
          // Без реєнтерабельності — deadlock!
          System.out.println("inner: монітор захоплено повторно (count=2)");
      }
  }` },
  
        { heading: "Synchronized на різних об'єктах" },
        { code: `public class MultiLockExample {
      private final Object lockA = new Object();
      private final Object lockB = new Object();
      private int dataA = 0, dataB = 0;
  
      // Блокує лише dataA, дозволяє паралельний доступ до dataB:
      public void updateA(int val) {
          synchronized (lockA) { dataA = val; }
      }
  
      // Блокує лише dataB:
      public void updateB(int val) {
          synchronized (lockB) { dataB = val; }
      }
  
      // НЕБЕЗПЕЧНО: різний порядок захоплення може призвести до deadlock!
      // Метод 1: synchronized(lockA) потім synchronized(lockB)
      // Метод 2: synchronized(lockB) потім synchronized(lockA) -- deadlock!
      // ПРАВИЛО: завжди захоплювати блокування в одному порядку!
      public void updateBoth(int a, int b) {
          synchronized (lockA) {
              synchronized (lockB) {
                  dataA = a; dataB = b;
              }
          }
      }
  }` },
  
        { heading: "ReentrantLock — явне блокування" },
        { text: "java.util.concurrent.locks.ReentrantLock надає більш гнучкий та потужний API ніж synchronized: tryLock() — спроба захопити без блокування; tryLock(timeout) — спроба з таймаутом (запобігає deadlock); lockInterruptibly() — можна перервати очікування через interrupt(); getQueueLength() — кількість потоків, що очікують; isHeldByCurrentThread() — чи утримується поточним потоком. ReentrantLock(true) — fair lock: забезпечує FIFO порядок отримання — усуває starvation, але знижує пропускну здатність." },
        { code: `ReentrantLock lock = new ReentrantLock(true); // fair
  
  public void safeUpdate() {
      lock.lock(); // блокуємось до отримання
      try {
          criticalSection();
      } finally {
          lock.unlock(); // ЗАВЖДИ у finally!
      }
  }
  
  public boolean tryUpdate(long timeoutMs) {
      try {
          if (lock.tryLock(timeoutMs, TimeUnit.MILLISECONDS)) {
              try {
                  criticalSection();
                  return true;
              } finally {
                  lock.unlock();
              }
          }
          return false; // не вдалося отримати
      } catch (InterruptedException e) {
          Thread.currentThread().interrupt();
          return false;
      }
  }` },
  
        { heading: "ReadWriteLock — оптимізація читань" },
        { code: `ReadWriteLock rwLock = new ReentrantReadWriteLock();
  Lock readLock  = rwLock.readLock();
  Lock writeLock = rwLock.writeLock();
  
  // Багато потоків можуть читати одночасно:
  public String getValue(String key) {
      readLock.lock();
      try { return cache.get(key); }
      finally { readLock.unlock(); }
  }
  
  // Лише один потік може писати (блокує всіх читачів):
  public void setValue(String key, String val) {
      writeLock.lock();
      try { cache.put(key, val); }
      finally { writeLock.unlock(); }
  }` },
        { text: "ReadWriteLock ефективний при READ-HEAVY навантаженні (конфігурація, кеш, довідники). При балансі читань і записів звичайний ReentrantLock може бути кращим через меншу складність." },
  
        { heading: "StampedLock (Java 8+) — оптимістичне читання" },
        { text: "StampedLock вводить optimistic read — спроба прочитати дані без захоплення блокування. Читач перевіряє stamp (мітку версії) до та після читання: якщо мітка не змінилася — запису не було, дані коректні. Значно ефективніший за ReadWriteLock при малій кількості записів." }
      ]
    },
  
    // ----------------------------------------------------------
    {
      id: 20,
      title: "20. Проблеми управління потоками: дедлок та інші небажані стани паралельної програми",
      answer: [
        { heading: "Deadlock — взаємне блокування" },
        { text: "Deadlock (взаємне блокування) — ситуація, коли два або більше потоки взаємно очікують один одного, утримуючи ресурси, потрібні іншим. Потоки застрягають назавжди. Умови Коффмана (необхідні та достатні для deadlock): 1) Взаємне виключення — ресурс може використовуватися лише одним потоком; 2) Утримання та очікування — потік утримує ресурс і очікує додаткового; 3) Неможливість примусового вилучення — ресурс не можна відібрати у потоку; 4) Кругове очікування — існує цикл очікування: A чекає B, B чекає C, C чекає A." },
        { code: `// КЛАСИЧНИЙ deadlock:
  Object lock1 = new Object();
  Object lock2 = new Object();
  
  Thread t1 = new Thread(() -> {
      synchronized (lock1) {
          System.out.println("T1: захопив lock1, чекає lock2");
          try { Thread.sleep(100); } catch (InterruptedException ignored) {}
          synchronized (lock2) { /* ... */ } // ЗАВИСАЄ — T2 тримає lock2
      }
  });
  Thread t2 = new Thread(() -> {
      synchronized (lock2) {                  // T2 захоплює lock2 першим
          System.out.println("T2: захопив lock2, чекає lock1");
          try { Thread.sleep(100); } catch (InterruptedException ignored) {}
          synchronized (lock1) { /* ... */ } // ЗАВИСАЄ — T1 тримає lock1
      }
  });` },
        { text: "Методи усунення deadlock: Lock Ordering — завжди захоплювати блокування в одному глобальному порядку (наприклад, за ідентифікатором об'єкта); tryLock з таймаутом — якщо не вдається отримати за час, відпустити вже захоплені та повторити; Lock-free алгоритми — використовувати AtomicReference та CAS замість mutex; мінімізація кількості одночасно захоплених блокувань." },
  
        { heading: "Starvation — голодування" },
        { text: "Starvation (голодування) — потік ніколи або дуже рідко отримує доступ до ресурсу через несправедливий планувальник або постійну конкуренцію. Потік живий (не заблокований), але не просувається. Причини: несправедливий (non-fair) synchronized — Java не гарантує порядку у черзі на монітор; низький пріоритет потоку при high-priority конкурентах; монополізація ресурсу довгими операціями. Рішення: ReentrantLock(true) — fair lock, FIFO порядок; збалансовані пріоритети; обмеження часу утримання блокування." },
  
        { heading: "Livelock" },
        { text: "Livelock — потоки активно змінюють стан у відповідь один на одного, але жоден не просувається до мети. Аналогія: двоє в коридорі намагаються поступитися дорогою — рухаються в один бік одночасно. Потоки не заблоковані (на відміну від deadlock), але марно витрачають CPU. Рішення: випадкова затримка перед повторною спробою (exponential backoff); глобальний арбітр." },
        { code: `// Livelock приклад:
  volatile boolean resource1Taken = false;
  volatile boolean resource2Taken = false;
  
  // Thread A:
  while (true) {
      resource1Taken = true;
      if (resource2Taken) { resource1Taken = false; continue; } // поступається
      useResources();
      break;
  }
  // Thread B аналогічно, але в іншому порядку — нескінченна взаємна поступка` },
  
        { heading: "Race Condition — стан гонитви" },
        { text: "Race condition — результат програми залежить від відносного порядку виконання потоків. Виникає при незахищеному доступі до спільних змінних. Check-then-act race: потік перевіряє умову (check), але до дії (act) умова змінюється іншим потоком. Read-modify-write race: i++ виглядає як одна операція, але є трьома (read, +1, write) — між ними може втрутитися інший потік. Виявлення: ThreadSanitizer, FindBugs/SpotBugs, Java Pathfinder (верифікація моделей)." },
  
        { heading: "Priority Inversion — інверсія пріоритетів" },
        { text: "Priority Inversion — потік з низьким пріоритетом утримує ресурс, потрібний потоку з високим пріоритетом. Між ними може виконуватися потік з середнім пріоритетом, ефективно блокуючи high-priority потік. Відома причина збою марсохода Pathfinder (1997). Вирішення: Priority Inheritance — ОС тимчасово підвищує пріоритет low-priority потоку до рівня high-priority, поки утримується спільний ресурс." },
  
        { heading: "Виявлення deadlock у JVM" },
        { code: `// Програмне виявлення через ThreadMXBean:
  ThreadMXBean mxBean = ManagementFactory.getThreadMXBean();
  long[] deadlockedIds = mxBean.findDeadlockedThreads();
  if (deadlockedIds != null) {
      ThreadInfo[] infos = mxBean.getThreadInfo(deadlockedIds, true, true);
      for (ThreadInfo info : infos) {
          System.err.println("DEADLOCK у потоці: " + info.getThreadName());
          System.err.println("Чекає на: " + info.getLockName());
      }
  }
  // Також: jstack <pid> у командному рядку показує deadlock` }
      ]
    },
  
    // ----------------------------------------------------------
    {
      id: 22,
      title: "22. Локери та управління потоками",
      answer: [
        { heading: "Lock інтерфейс" },
        { text: "java.util.concurrent.locks.Lock — базовий інтерфейс для явних блокувань у JUC. Методи: lock() — захоплює блокування, блокуючи потік якщо воно зайняте; unlock() — звільняє (завжди у finally!); tryLock() — спроба без блокування, повертає boolean; tryLock(long time, TimeUnit unit) — з таймаутом; lockInterruptibly() — захоплення з можливістю переривання; newCondition() — створює Condition (аналог Object.wait/notify) прив'язаний до цього локера." },
  
        { heading: "ReentrantLock — детальний розбір" },
        { code: `ReentrantLock lock = new ReentrantLock();
  Condition notEmpty = lock.newCondition();
  Condition notFull  = lock.newCondition();
  final int CAPACITY = 10;
  Queue<Integer> buffer = new ArrayDeque<>();
  
  // Producer:
  public void produce(int item) throws InterruptedException {
      lock.lock();
      try {
          while (buffer.size() == CAPACITY)
              notFull.await(); // чекає на notFull.signal()
          buffer.offer(item);
          notEmpty.signal(); // сигналізуємо consumer'у
      } finally { lock.unlock(); }
  }
  
  // Consumer:
  public int consume() throws InterruptedException {
      lock.lock();
      try {
          while (buffer.isEmpty())
              notEmpty.await();
          int item = buffer.poll();
          notFull.signal();
          return item;
      } finally { lock.unlock(); }
  }` },
        { text: "Перевага Condition над Object.wait/notify: можна мати кілька умов очікування на одному об'єкті (notEmpty та notFull окремо). З synchronized довелося б використовувати notifyAll(), що будить усі потоки навіть коли тільки одна умова змінилася." },
  
        { heading: "StampedLock" },
        { code: `StampedLock sl = new StampedLock();
  double x, y;
  
  // Optimistic read — найшвидший, без блокування:
  public double distanceFromOrigin() {
      long stamp = sl.tryOptimisticRead();
      double cx = x, cy = y;
      if (!sl.validate(stamp)) { // хтось писав — перечитуємо з блокуванням
          stamp = sl.readLock();
          try { cx = x; cy = y; }
          finally { sl.unlockRead(stamp); }
      }
      return Math.hypot(cx, cy);
  }
  
  // Write lock:
  public void move(double dx, double dy) {
      long stamp = sl.writeLock();
      try { x += dx; y += dy; }
      finally { sl.unlockWrite(stamp); }
  }` },
  
        { heading: "AbstractQueuedSynchronizer (AQS)" },
        { text: "AQS — серце більшості синхронізаторів JUC. ReentrantLock, Semaphore, CountDownLatch, CyclicBarrier, ReadWriteLock — всі базуються на AQS. AQS управляє чергою очікуючих потоків (CLH-черга на основі linked list) та атомарним int state. Підклас перевизначає tryAcquire()/tryRelease() (ексклюзивний режим) або tryAcquireShared()/tryReleaseShared() (спільний режим) для визначення логіки блокування." },
  
        { heading: "Управління потоками через ScheduledExecutorService" },
        { code: `ScheduledExecutorService scheduler =
      Executors.newScheduledThreadPool(2);
  
  // Одноразове відкладене виконання:
  scheduler.schedule(
      () -> System.out.println("Виконано через 5 сек"),
      5, TimeUnit.SECONDS
  );
  
  // Повторне виконання з фіксованим ІНТЕРВАЛОМ між запусками:
  scheduler.scheduleAtFixedRate(
      () -> collectMetrics(),
      0, 10, TimeUnit.SECONDS // initial delay=0, period=10s
  );
  
  // Повторне виконання з фіксованою ЗАТРИМКОЮ після завершення:
  scheduler.scheduleWithFixedDelay(
      () -> processQueue(),
      1, 5, TimeUnit.SECONDS // після кожного завершення чекати 5с
  );
  
  // Граційне завершення:
  scheduler.shutdown();
  if (!scheduler.awaitTermination(30, TimeUnit.SECONDS))
      scheduler.shutdownNow();` },
  
        { heading: "CompletableFuture як управління ланцюжками задач" },
        { text: "CompletableFuture (Java 8+) дозволяє будувати non-blocking ланцюжки асинхронних задач: thenApply (трансформація результату у тому ж потоці); thenApplyAsync (у пулі потоків); thenCombine (поєднання двох futures); thenCompose (flatMap для futures); exceptionally / handle (обробка помилок); allOf / anyOf (очікування кількох futures)." }
      ]
    },
  
    // ----------------------------------------------------------
    {
      id: 24,
      title: "24. Інтерфейс Executor та бібліотечні класи, які його реалізують",
      answer: [
        { heading: "Ієрархія Executor Framework" },
        { text: "Executor Framework — архітектура для відокремлення задач (що виконувати) від механізму виконання (де та коли виконувати). Ієрархія інтерфейсів: Executor (один метод execute(Runnable)) → ExecutorService (submit, shutdown, invokeAll, invokeAny) → ScheduledExecutorService (schedule, scheduleAtFixedRate). Реалізації: ThreadPoolExecutor (основа всього), ScheduledThreadPoolExecutor, ForkJoinPool." },
  
        { heading: "Інтерфейс Executor" },
        { code: `// Найпростіший інтерфейс: один метод
  @FunctionalInterface
  public interface Executor {
      void execute(Runnable command);
  }
  
  // Можна реалізувати по-різному:
  Executor sameThread = Runnable::run;          // у поточному потоці
  Executor newThread  = r -> new Thread(r).start(); // у новому потоці
  Executor poolThread = Executors.newFixedThreadPool(4); // у пулі` },
  
        { heading: "ExecutorService — розширений інтерфейс" },
        { text: "ExecutorService додає до Executor: submit(Callable<T>) → Future<T> — задача з результатом; submit(Runnable) → Future<?> — задача без результату; invokeAll(Collection<Callable<T>>) — запуск усіх та отримання списку Future; invokeAny(Collection<Callable<T>>) — повертає результат першої успішної задачі; shutdown() — більше не приймає задачі, дочікується завершення поточних; shutdownNow() — намагається зупинити поточні задачі (interrupt), повертає незапущені; awaitTermination(timeout) — чекає завершення після shutdown." },
        { code: `ExecutorService exec = Executors.newFixedThreadPool(4);
  
  // submit() + Future:
  List<Future<Double>> futures = new ArrayList<>();
  for (int i = 0; i < 100; i++) {
      final int idx = i;
      futures.add(exec.submit(() -> heavyCalc(idx)));
  }
  // Збираємо результати:
  double sum = 0;
  for (Future<Double> f : futures) sum += f.get();
  
  // invokeAll() — зручніше:
  List<Callable<Double>> tasks = IntStream.range(0, 100)
      .mapToObj(i -> (Callable<Double>) () -> heavyCalc(i))
      .collect(Collectors.toList());
  List<Future<Double>> results = exec.invokeAll(tasks);
  
  exec.shutdown();
  exec.awaitTermination(60, TimeUnit.SECONDS);` },
  
        { heading: "ThreadPoolExecutor — деталі" },
        { text: "ThreadPoolExecutor(corePoolSize, maxPoolSize, keepAlive, TimeUnit, BlockingQueue, ThreadFactory, RejectedExecutionHandler). Логіка: при надходженні задачі якщо живих потоків < corePoolSize — створюємо новий потік; якщо >= core і черга не повна — кладемо в чергу; якщо черга повна і потоків < maxPoolSize — створюємо додатковий потік; якщо потоків == maxPoolSize і черга повна — викликаємо RejectedExecutionHandler." },
        { code: `ThreadPoolExecutor pool = new ThreadPoolExecutor(
      4,                              // corePoolSize
      16,                             // maximumPoolSize
      60L, TimeUnit.SECONDS,          // keepAliveTime
      new LinkedBlockingQueue<>(200), // workQueue
      Executors.defaultThreadFactory(),
      new ThreadPoolExecutor.CallerRunsPolicy() // якщо черга переповнена
  );
  
  // CallerRunsPolicy — задача виконується у потоці-відправнику:
  // природне throttling без виключень
  
  // Моніторинг:
  System.out.println("Active: " + pool.getActiveCount());
  System.out.println("Queue size: " + pool.getQueue().size());
  System.out.println("Completed: " + pool.getCompletedTaskCount());` },
  
        { heading: "Фабричні методи Executors" },
        { text: "Executors.newFixedThreadPool(n) — n потоків, необмежена черга LinkedBlockingQueue. Для CPU-bound задач: n = кількість ядер. Executors.newCachedThreadPool() — потоки на вимогу, виживають 60 с. Для короткочасних IO-bound задач. Executors.newSingleThreadExecutor() — один потік, гарантована послідовність. Executors.newScheduledThreadPool(n) — планувальник задач. Executors.newWorkStealingPool(n) — обгортка над ForkJoinPool для незалежних задач." },
  
        { heading: "ForkJoinPool як ExecutorService" },
        { text: "ForkJoinPool реалізує ExecutorService з work-stealing планувальником. ForkJoinPool.commonPool() — спільний пул з parallelism = кількість ядер - 1. Можна використовувати як звичайний ExecutorService (submit(Callable)), але ефективніший з RecursiveTask/RecursiveAction через fork/join механізм. Java parallelStream() використовує commonPool() всередині." }
      ]
    },
  
    // ----------------------------------------------------------
    {
      id: 26,
      title: "26. Методи моделювання паралельних обчислень",
      answer: [
        { heading: "Навіщо моделювати паралельні обчислення" },
        { text: "Розробка паралельних програм складніша за послідовну: поведінка недетермінована, помилки (race conditions, deadlock) важко відтворити, а відлагодження змінює часові характеристики (Heisenbug). Моделювання дозволяє: верифікувати коректність алгоритму до реалізації; оцінити теоретичну продуктивність; виявити потенційні deadlock та race condition аналітично; спростити аналіз через абстракцію від деталей реалізації." },
  
        { heading: "Граф DAG (операції-операнди)" },
        { text: "Направлений ациклічний граф (DAG) операцій: вузли — обчислювальні операції; ребра — залежності за даними (ребро A→B: B потребує результату A). DAG дозволяє визначити: які операції можна виконати паралельно (відсутність ребер між ними); критичний шлях — найдовший шлях від вхідних до вихідних вузлів = мінімальний час виконання при нескінченних ресурсах; необхідна кількість паралельних одиниць — максимальна ширина DAG (найбільша кількість вузлів на одному 'рівні')." },
  
        { heading: "Мережі Петрі" },
        { text: "Мережа Петрі (Petri Net) — математична модель для опису розподілених систем. Складові: місця (places, кола) — стани або буфери; переходи (transitions, прямокутники) — події або операції; дуги (arcs) — зв'язки між місцями та переходами; маркери (tokens) — наявність ресурсів або дозволів у місцях. Перехід спрацьовує (fires), коли у кожному вхідному місці є щонайменше по одному маркеру. При спрацюванні маркери знімаються з вхідних місць та додаються у вихідні." },
        { text: "Мережі Петрі моделюють: синхронізацію (AND-join: потік продовжується коли всі вхідні місця мають маркери); вибір (OR-split: маркер іде в одне з кількох вихідних місць); паралелізм (AND-split: маркери додаються у кілька вихідних місць одночасно); конкуренцію за ресурси (місце з одним маркером — ексклюзивний доступ). Аналіз мереж Петрі: reachability (чи досяжний певний стан); boundedness (чи не переповнюється місце); liveness (чи можливе deadlock — існування 'мертвого' стану)." },
  
        { heading: "BSP (Bulk Synchronous Parallel)" },
        { text: "BSP модель (Valiant, 1990) описує паралельні обчислення як послідовність супер-кроків (supersteps). Кожен супер-крок: локальні обчислення (кожен процесор незалежно); комунікація (надсилання повідомлень); бар'єрна синхронізація. Параметри BSP: p — кількість процесорів; l — вартість бар'єру (latency); g — відношення часу комунікації до часу обчислення. Вартість супер-кроку: max(w_i) + g*max(h_i) + l. BSP гарантує детермінованість та композиційність програм." },
  
        { heading: "PRAM (Parallel Random Access Machine)" },
        { text: "PRAM — теоретична модель для аналізу складності паралельних алгоритмів. p ідентичних процесорів із спільною пам'яттю без затримки доступу; за один крок всі процесори виконують одну операцію. Варіанти за можливістю конфліктів: EREW (Exclusive Read Exclusive Write) — найобмеженіша; CREW (Concurrent Read Exclusive Write) — паралельне читання дозволено; CRCW (Concurrent Read Concurrent Write) — найпотужніша. PRAM є теоретичним інструментом; реальні машини є EREW або CREW через когерентність кешу." },
  
        { heading: "Логічний час та відношення happens-before (Лампорт)" },
        { text: "У розподілених системах немає глобального годинника. Лампорт (1978) ввів логічний час: кожен процес має лічильник; при настанні події лічильник збільшується на 1; при надсиланні повідомлення — разом з повідомленням передається поточне значення лічильника; при отриманні — лічильник отримувача = max(local, received) + 1. Це дозволяє частково впорядкувати події: якщо a → b (a передує b), то timestamp(a) < timestamp(b). Вектор-годинники (Fidge, Mattern) забезпечують точне відображення happens-before." }
      ]
    },
  
    // ----------------------------------------------------------
    {
      id: 28,
      title: "28. Моделювання паралельних програм мережею Петрі",
      answer: [
        { heading: "Базова структура мережі Петрі" },
        { text: "Мережа Петрі N = (P, T, F, W, M₀), де P — скінченна множина місць; T — скінченна множина переходів; F ⊆ (P×T) ∪ (T×P) — функція потоку (дуги); W: F → N⁺ — вагова функція (кількість маркерів на дузі, за замовчуванням 1); M₀: P → N — початкове маркування. Перехід t є дозволеним (enabled), якщо для кожного вхідного місця p: M(p) ≥ W(p,t). При спрацюванні: M'(p) = M(p) - W(p,t) для вхідних місць; M'(p) = M(p) + W(t,p) для вихідних місць." },
  
        { heading: "Моделювання паттерну Producer-Consumer" },
        { text: "Місця: IDLE_PRODUCER (виробник вільний); PRODUCING (виробник виробляє); BUFFER (буфер, ємність = N маркерів); IDLE_CONSUMER (споживач вільний); CONSUMING (споживач споживає). Переходи: t1 (start_produce): вхід IDLE_PRODUCER → вихід PRODUCING; t2 (finish_produce): вхід PRODUCING → виходи BUFFER + IDLE_PRODUCER; t3 (start_consume): вхід BUFFER + IDLE_CONSUMER → вихід CONSUMING; t4 (finish_consume): вхід CONSUMING → IDLE_CONSUMER. Обмеження буфера: додати місце BUFFER_SPACE з N маркерами та дугу від нього до t2 і від t4 до нього. Тепер t2 не спрацює, якщо буфер повний." },
  
        { heading: "Моделювання взаємного виключення (mutex)" },
        { text: "Місця: MUTEX (1 маркер = вільний, 0 = зайнятий); READY_1, CRITICAL_1, DONE_1 (стани потоку 1); READY_2, CRITICAL_2, DONE_2 (стани потоку 2). Переходи: enter_1: вхід READY_1 + MUTEX → вихід CRITICAL_1; exit_1: вхід CRITICAL_1 → виходи DONE_1 + MUTEX; аналогічно для потоку 2. Маркер у MUTEX — 'ключ' від критичної секції. Обидва потоки не можуть одночасно знаходитися у CRITICAL, оскільки маркер MUTEX один. Deadlock неможливий (mutex завжди повертається)." },
  
        { heading: "Аналіз мережі Петрі" },
        { text: "Граф досяжності (reachability graph): вершини — всі досяжні маркування; ребра — спрацьовування переходів. Якщо граф скінченний — мережа обмежена (bounded). Аналіз властивостей: Safety (безпечність): місце обмежене якщо кількість маркерів у ньому ніколи не перевищує деяке k (k-bounded). Deadlock-freeness: немає досяжного маркування без жодного дозволеного переходу. Liveness: перехід t є live якщо з будь-якого досяжного маркування можна дійти до маркування де t дозволений. Fairness: кожен дозволений перехід рано чи пізно спрацьовує." },
  
        { heading: "Кольорові мережі Петрі (CPN)" },
        { text: "Базові мережі Петрі мають лише 'анонімні' маркери. Кольорові мережі Петрі (CPN, Kurt Jensen) розширюють модель: маркери мають типи (кольори) — цілі числа, рядки, кортежі; дуги мають вирази для вибору маркерів; переходи мають охоронні умови. CPN значно потужніші — можна моделювати конкретні значення даних, параметризовані протоколи. Інструментарій: CPN Tools (безкоштовний), GreatSPN, TINA." },
  
        { heading: "Практичне застосування для паралельних Java програм" },
        { text: "Мережа Петрі для ForkJoin алгоритму: місце TASK_POOL — початковий пул задач; перехід FORK — розбиває задачу на підзадачі (один маркер → два); місця SUBTASK_1, SUBTASK_2 — підзадачі; переходи COMPUTE_1, COMPUTE_2 — виконання; місця RESULT_1, RESULT_2 — результати; перехід JOIN — вимагає обидва результати (два маркери → один). Це точно моделює sematics RecursiveTask: join() блокується поки обидві гілки не завершені, що відображається необхідністю двох маркерів для спрацьовування JOIN." }
      ]
    },
  
    // ----------------------------------------------------------
    {
      id: 30,
      title: "30. Алгоритми паралельного сумування та оцінка їх ефективності",
      answer: [
        { heading: "Послідовне сумування та його обмеження" },
        { text: "Послідовна сума масиву з N елементів: O(N) операцій, O(N) часу, один процесор. При N = 10⁸ та 1 нс на операцію: 100 мс. Мета паралелізації: при p процесорах зменшити час до O(N/p + log p) — паралельна частина + накладні витрати злиття." },
  
        { heading: "Деревоподібна (tree-based) паралельна сума" },
        { text: "Алгоритм: кожен рядок дерева виконується паралельно. Рівень 0: N чисел. Рівень 1: N/2 потоків, кожен сумує пару → N/2 проміжних результатів. Рівень 2: N/4 потоків → N/4 результатів. ... Рівень log₂N: 1 результат — фінальна сума. Загальна кількість операцій: N-1 (оптимально). Час при N/2 процесорах: O(log N) кроків. Прискорення: S = N / log N. Ефективність: E = N / (N/2 * log N) = 2/log N → 0 при N→∞ (не масштабується без оптимізацій)." },
  
        { heading: "Реалізація з ForkJoinPool (RecursiveTask)" },
        { code: `public class ParallelSum extends RecursiveTask<Long> {
      static final int THRESHOLD = 10_000;
      private final long[] array;
      private final int from, to;
  
      public ParallelSum(long[] array, int from, int to) {
          this.array = array; this.from = from; this.to = to;
      }
  
      @Override
      protected Long compute() {
          if (to - from <= THRESHOLD) {
              long sum = 0;
              for (int i = from; i < to; i++) sum += array[i];
              return sum;
          }
          int mid = (from + to) >>> 1;
          ParallelSum left  = new ParallelSum(array, from, mid);
          ParallelSum right = new ParallelSum(array, mid,  to);
          left.fork();                           // async запуск лівої
          long rightResult = right.compute();    // sync права (у поточному потоці)
          long leftResult  = left.join();        // чекаємо лівої
          return leftResult + rightResult;
      }
  
      public static long sum(long[] array) {
          return ForkJoinPool.commonPool()
              .invoke(new ParallelSum(array, 0, array.length));
      }
  }` },
  
        { heading: "Паралельне prefix-sum (Scan)" },
        { text: "Prefix Sum (Scan) — обчислення масиву часткових сум: output[i] = input[0] + input[1] + ... + input[i]. Здається послідовним, але має ефективний паралельний алгоритм. Алгоритм Blelloch (1990): Фаза 1 (Up-sweep / Reduce): будуємо дерево часткових сум знизу вгору — O(N) роботи, O(log N) глибина. Фаза 2 (Down-sweep): розповсюджуємо суми зверху вниз — O(N) роботи, O(log N) глибина. Разом: O(N) роботи, O(log N) паралельного часу. Застосування: квантильна нормалізація; паралельний QuickSort (partition); стиснення даних (run-length encoding)." },
  
        { heading: "Паралельна сума через Stream API" },
        { code: `// Java 8 parallel stream — автоматична паралелізація:
  long[] array = LongStream.rangeClosed(1, 1_000_000).toArray();
  
  long seqSum = LongStream.of(array).sum();          // послідовно
  long parSum = LongStream.of(array).parallel().sum(); // паралельно
  
  // IntStream, LongStream, DoubleStream мають оптимізовані
  // спеціалізовані reduction операції:
  double avg = Arrays.stream(array).parallel().average().getAsDouble();
  long   max = Arrays.stream(array).parallel().max().getAsLong();
  
  // Загальна форма parallel reduce:
  int result = IntStream.range(0, N)
      .parallel()
      .reduce(0, (a, b) -> a + b); // identity + combiner` },
  
        { heading: "Оцінка ефективності паралельного сумування" },
        { text: "При N елементах та p процесорах оптимальна стратегія: кожен процесор підраховує суму N/p елементів → час T_local = N/p; потім деревоподібне злиття p часткових сум → час T_merge = log₂p. Загальний час: Tₚ = N/p + log₂p. Прискорення: Sₚ = N / (N/p + log₂p). При N >> p*log₂p: Sₚ ≈ p (ідеальне прискорення). Ефективність: Eₚ ≈ 1 - p*log₂p/N. Для N = 10⁶, p = 8: E₈ ≈ 1 - 24/10⁶ ≈ 99.997% — практично ідеально." }
      ]
    },
  
    // ----------------------------------------------------------
    {
      id: 32,
      title: "32. Експериментальне дослідження ефективності паралельних обчислень",
      answer: [
        { heading: "Методологія бенчмаркінгу" },
        { text: "Вимірювання продуктивності паралельних програм вимагає особливої обережності. Типові помилки: вимірювання лише першого запуску (JVM не прогріта, JIT не скомпілював код); вимірювання з різними вхідними даними; змішування CPU time та wall clock time. Правильна методологія: прогрів (warmup): запустити алгоритм 5–20 разів без вимірювань для JIT компіляції; множинні вимірювання: мінімум 10–30 повторень; статистика: медіана або середнє без outliers; ізоляція: вимкнути GC паузи (G1GC, ZGC), виключні ядра (CPU affinity); однакові умови для всіх конфігурацій." },
  
        { heading: "JMH — Java Microbenchmark Harness" },
        { text: "JMH (Oracle/OpenJDK) є стандартом для точного вимірювання продуктивності Java коду. Коректно обробляє JIT warmup, dead code elimination, constant folding." },
        { code: `@BenchmarkMode(Mode.AverageTime)
  @OutputTimeUnit(TimeUnit.MILLISECONDS)
  @State(Scope.Thread)
  @Warmup(iterations = 5, time = 1, timeUnit = TimeUnit.SECONDS)
  @Measurement(iterations = 10, time = 2, timeUnit = TimeUnit.SECONDS)
  @Fork(value = 2) // 2 окремих JVM-процеси
  public class MatrixMultiplyBenchmark {
  
      @Param({"256", "512", "1024"})
      int size;
  
      double[][] A, B;
  
      @Setup(Level.Trial)
      public void setUp() {
          A = randomMatrix(size);
          B = randomMatrix(size);
      }
  
      @Benchmark
      public double[][] sequential() {
          return MatrixMul.sequential(A, B);
      }
  
      @Benchmark
      public double[][] parallel4() {
          return MatrixMul.parallel(A, B, 4);
      }
  
      @Benchmark
      public double[][] forkjoin() {
          return MatrixMul.forkJoin(A, B);
      }
  }` },
  
        { heading: "Власна реалізація вимірювань" },
        { code: `public class BenchmarkRunner {
      public static BenchmarkResult measure(
              Callable<Object> task, int warmup, int runs) throws Exception {
  
          // Прогрів:
          for (int i = 0; i < warmup; i++) task.call();
  
          // Вимірювання:
          long[] times = new long[runs];
          for (int i = 0; i < runs; i++) {
              // Підказка GC перед вимірюванням:
              System.gc();
              Thread.sleep(50);
              long start = System.nanoTime();
              task.call();
              times[i] = System.nanoTime() - start;
          }
          Arrays.sort(times);
          long median = times[runs / 2];
          long mean   = LongStream.of(times).sum() / runs;
          long stdDev = (long) Math.sqrt(
              LongStream.of(times).mapToLong(t -> (t-mean)*(t-mean)).sum() / runs
          );
          return new BenchmarkResult(median, mean, stdDev);
      }
  }` },
  
        { heading: "Побудова графіків масштабованості" },
        { code: `public class ScalabilityStudy {
      public static void main(String[] args) throws Exception {
          int N = 1000;
          double[][] A = generateMatrix(N), B = generateMatrix(N);
  
          System.out.printf("%-10s %-12s %-14s %-14s%n",
              "Потоки", "Час (мс)", "Прискорення", "Ефективність");
  
          long T1 = 0;
          int maxP = Runtime.getRuntime().availableProcessors();
          for (int p = 1; p <= maxP; p++) {
              final int threads = p;
              long time = BenchmarkRunner.measure(
                  () -> MatrixMul.parallel(A, B, threads), 3, 10).median;
              if (p == 1) T1 = time;
  
              double speedup = (double) T1 / time;
              double efficiency = speedup / p;
              double amdahl = 1.0 / (0.05 + 0.95 / p); // assume 5% serial
              System.out.printf("%-10d %-12.2f %-14.3f %-14.2f%% (Amdahl: %.3f)%n",
                  p, time / 1e6, speedup, efficiency * 100, amdahl);
          }
      }
  }` },
  
        { heading: "Аналіз результатів та типові патерни" },
        { text: "Лінійне прискорення (S ≈ p) — зазвичай досягається при великих задачах без синхронізацій та рівномірному розподілі. Закон Амдала (плато) — прискорення виходить на межу 1/f через послідовну частину f. Деградація при великих p — overhead від синхронізації, false sharing, load imbalance. Суперлінійне прискорення (S > p) — кешові ефекти: при p процесорах задача вміщується в сумарний кеш. Cache effect: N=2000 матриця = 32 MB, не вміщується в L3 одного ядра, але при p=8 кожен тримає 4 MB — вміщується в L3. Профілювання: async-profiler, Intel VTune — виявляють: cache miss; false sharing; spinlock contention; imbalanced work distribution." }
      ]
    },
  
    // ----------------------------------------------------------
    {
      id: 34,
      title: "34. Проектування паралельних програм",
      answer: [
        { heading: "Методологія PCAM (Ian Foster)" },
        { text: "Ian Foster запропонував методологію PCAM для систематичного проектування паралельних програм: Partitioning (розбиття) — розбити задачу на дрібні підзадачі, максимізуючи паралелізм; Communication (комунікація) — визначити необхідний обмін даними між підзадачами; Agglomeration (агрегація) — об'єднати підзадачі для зменшення накладних витрат та покращення locality; Mapping (відображення) — призначити агреговані задачі процесорам." },
  
        { heading: "Крок 1: Partitioning" },
        { text: "Мета: виявити максимальний ступінь паралелізму. Декомпозиція за даними: розбити вхідні дані на підмножини, застосувати однакову операцію до кожної. Приклад: рядки матриці → кожен рядок незалежно обчислюється. Декомпозиція за задачами: розбити алгоритм на незалежні функціональні блоки. Приклад: читання файлу + парсинг + збереження можуть виконуватися конвеєрно. Правила хорошого розбиття: підзадачі однакового розміру (рівномірне навантаження); кількість підзадач >> кількість процесорів (гнучкість планування); мінімальна залежність між підзадачами." },
  
        { heading: "Крок 2: Communication" },
        { text: "Визначити, які дані потрібно передавати між підзадачами. Local communication: кожна підзадача спілкується з обмеженим числом сусідів (стрічковий алгоритм — обмін граничними рядками). Global communication: всі підзадачі спілкуються між собою (MPI_AllReduce). Патерни комунікацій: point-to-point (MPI_Send/Recv); broadcast (один → всі); scatter/gather (root ↔ всі); alltoall (кожен з кожним)." },
  
        { heading: "Крок 3: Agglomeration (Конкретизація)" },
        { text: "Мета: зменшити накладні витрати, зберігаючи масштабованість. Об'єднати дрібні задачі у більші ('зерна' обчислень). Для задачі з N підзадач та p процесорів: об'єднати N/p підзадач в одну (стрічку/блок). Зберегти flexibility: залишити можливість налаштування granularity. У ForkJoin: THRESHOLD — мінімальний розмір для рекурсивного розбиття, нижче якого виконуємо послідовно." },
  
        { heading: "Крок 4: Mapping" },
        { text: "Призначення задач процесорам. Статичне відображення: задачі розподіляються до початку виконання. Підходить для рівномірних задач (стрічковий алгоритм). Динамічне відображення (load balancing): задачі розподіляються під час виконання. Підходить для нерівномірних задач (обхід дерева). Work-stealing у ForkJoin: незайнятий потік краде задачі у зайнятого — автоматичне динамічне балансування." },
  
        { heading: "Шаблони паралельного проектування" },
        { text: "Master-Worker: Master розподіляє роботу, Workers виконують та повертають результати. FixedThreadPool + task queue — класична реалізація. Pipeline: обробка даних через послідовність стадій, паралельно для різних порцій даних. BlockingQueue між стадіями. Map-Reduce: Map застосовує функцію до незалежних підмножин, Reduce об'єднує результати. ParallelStream у Java. Fork-Join: рекурсивне розбиття на підзадачі з деревоподібним збиранням. ForkJoinPool. Speculative Execution: паралельний запуск кількох варіантів, перший результат перемагає. CompletableFuture.anyOf()." },
  
        { heading: "Балансування навантаження" },
        { code: `// Динамічне балансування через task queue:
  ExecutorService pool = Executors.newFixedThreadPool(
      Runtime.getRuntime().availableProcessors()
  );
  
  // Кожна задача (chunk) — один рядок матриці:
  // Дрібніші задачі → краще балансування нерівномірних операцій
  List<Future<double[]>> results = new ArrayList<>();
  for (int i = 0; i < N; i++) {
      final int row = i;
      results.add(pool.submit(() -> computeRow(A, B, row, N)));
  }
  
  // Збираємо результати в порядку завершення:
  double[][] C = new double[N][];
  for (int i = 0; i < N; i++) {
      C[i] = results.get(i).get();
  }
  pool.shutdown();` }
      ]
    },
  
    // ----------------------------------------------------------
    {
      id: 36,
      title: "36. Моделі пам'яті паралельних обчислень",
      answer: [
        { heading: "Проблема моделі пам'яті" },
        { text: "Сучасні процесори та компілятори виконують оптимізації (reordering, caching, buffering), що порушують інтуїтивне розуміння 'що коли видно'. Модель пам'яті визначає: які значення може прочитати потік/процес, якщо інший потік записав значення. Без чіткої специфікації неможливо писати коректні паралельні програми. Java Memory Model (JMM) — одна з перших офіційних специфікацій моделі пам'яті для мов загального призначення." },
  
        { heading: "Переупорядкування (Reordering)" },
        { text: "Три рівні переупорядкування: Компілятор (JIT): може переставляти незалежні інструкції, усувати 'непотрібні' операції. Процесор: out-of-order execution, store buffers, load forwarding. Кеш-пам'ять: запис у store buffer ще не видимий іншим ядрам. Для однопоточної програми переупорядкування не спостерігається — результат такий самий як при послідовному виконанні (as-if-serial semantics). Для багатопоточних — реупорядкування між потоками є джерелом проблем." },
  
        { heading: "Java Memory Model (JMM) та Happens-Before" },
        { text: "JMM (JSR-133, JDK 5) базується на happens-before (hb) відношенні. Дія A hb B означає: ефекти A видимі B. Правила hb в Java: Program order: у межах одного потоку всі дії hb наступні; Monitor unlock hb наступний lock того ж монітора; Volatile write hb наступне читання того ж поля; Thread.start() hb перша дія запущеного потоку; Thread.join() hb: всі дії потоку hb успішне повернення з join(); Transitivity: A hb B, B hb C → A hb C. Якщо немає hb-відношення між записом і читанням — читання може побачити будь-яке значення (включаючи 'стале')." },
  
        { heading: "volatile — гарантії JMM" },
        { code: `// volatile ГАРАНТУЄ:
  // 1. Видимість: запис volatile одразу видимий читачу
  // 2. Заборону reordering навколо volatile операцій (memory barrier)
  // volatile НЕ ГАРАНТУЄ атомарність compound операцій:
  
  class Wrong {
      volatile int count = 0;
      void increment() { count++; } // НЕ АТОМАРНО! read-modify-write
  }
  
  class Correct {
      AtomicInteger count = new AtomicInteger(0);
      void increment() { count.incrementAndGet(); } // атомарно
  }
  
  // Типове правильне використання volatile:
  class Flag {
      private volatile boolean shutdown = false;
      public void stop() { shutdown = true; }  // write hb read
      public void run() {
          while (!shutdown) { doWork(); }      // читає актуальне значення
      }
  }` },
  
        { heading: "Модель пам'яті процесора: x86 vs ARM" },
        { text: "x86 TSO (Total Store Order) — відносно сувора модель: читання не переставляються з читаннями; записи не переставляються із записами; читання не переставляються із записами; але запис може бути затримано у store buffer (інші ядра не бачать відразу). Тому на x86 volatile потребує лише store-load бар'єру. ARM/AArch64 — слабша модель пам'яті (weak memory model): набагато більше дозволених переупорядкувань; потребує явних DMB (data memory barrier) інструкцій. Тому синхронізаційний код, що 'працює' на x86, може ламатися на ARM." },
  
        { heading: "Модель пам'яті MPI" },
        { text: "У MPI немає спільної пам'яті між процесами — передача повідомлень є єдиним способом обміну. MPI_Send/MPI_Recv встановлюють happens-before: завершення Send hb початок обробки Recv. Однобічні комунікації (RMA: MPI_Put, MPI_Get) вимагають явних синхронізацій: MPI_Win_fence або MPI_Win_lock/unlock для встановлення видимості." }
      ]
    },
  
    // ----------------------------------------------------------
    {
      id: 38,
      title: "38. Стандарт Message Passing Interface (MPI)",
      answer: [
        { heading: "Що таке MPI" },
        { text: "MPI (Message Passing Interface) — відкритий стандарт для передачі повідомлень у паралельних і розподілених системах. Розроблений MPI Forum — консорціумом університетів, лабораторій та компаній. Актуальна версія MPI-4.1 (2023). MPI не є бібліотекою — це специфікація інтерфейсу. Реалізації: OpenMPI (найпоширеніша, open source); MPICH (ANL/MSU); Intel MPI; Cray MPICH; MPJ Express (Java). MPI описує понад 430 функцій, більшість програм використовує лише ~20–30." },
  
        { heading: "Основні концепції MPI" },
        { text: "Процеси та ранги: MPI-програма запускається як N незалежних процесів (mpirun -np N). Кожен процес має ранг (rank) від 0 до N-1 у рамках комунікатора. Ранг є унікальним ідентифікатором. Комунікатор (Communicator): набір процесів з спільним контекстом комунікацій. MPI_COMM_WORLD — стандартний комунікатор усіх процесів. Типи даних: MPI_INT, MPI_DOUBLE, MPI_CHAR тощо — платформонезалежне представлення типів. Теги (Tags): цілочисельні мітки повідомлень для розрізнення різних типів обміну між тими ж процесами." },
  
        { heading: "Структура MPI-програми на Java (MPJ Express)" },
        { code: `import mpi.*;
  
  public class HelloMPI {
      public static void main(String[] args) throws MPIException {
          MPI.Init(args);
  
          int rank = MPI.COMM_WORLD.Rank();
          int size = MPI.COMM_WORLD.Size();
          String host = MPI.Get_processor_name();
  
          System.out.printf("Процес %d з %d на %s%n", rank, size, host);
  
          // Point-to-point: процес 0 надсилає "ping" процесу 1:
          if (rank == 0) {
              int[] msg = {42};
              MPI.COMM_WORLD.Send(msg, 0, 1, MPI.INT, 1, 99);
              System.out.println("Процес 0: надіслав 42");
          } else if (rank == 1) {
              int[] buf = new int[1];
              Status st = MPI.COMM_WORLD.Recv(buf, 0, 1, MPI.INT,
                  MPI.ANY_SOURCE, 99);
              System.out.println("Процес 1: отримав " + buf[0]
                  + " від " + st.source);
          }
  
          MPI.Finalize();
      }
  }
  // mpirun -np 2 java -cp .:mpj.jar HelloMPI` },
  
        { heading: "Point-to-point комунікації" },
        { text: "MPI_Send(buf, count, datatype, dest, tag, comm) — стандартне надсилання. Повертається після того, як буфер можна безпечно перевикористати (або скопійовано в системний буфер, або доставлено). MPI_Recv(buf, count, datatype, source, tag, comm, status) — отримання. Блокується до повного отримання. MPI_Sendrecv — атомарний надіслати-і-отримати. Запобігає deadlock при кільцевих обмінах. MPI_Isend / MPI_Irecv — неблокуючі версії. Повертають Request-об'єкт. MPI_Wait(request) — чекати завершення." },
  
        { heading: "Колективні операції" },
        { text: "MPI_Bcast(buf, count, type, root, comm) — broadcast: root надсилає всім. MPI_Scatter(sendbuf, sendcount, ..., recvbuf, recvcount, ..., root, comm) — root розподіляє рівні частини масиву всім процесам. MPI_Gather — зворотній Scatter. MPI_Allgather — Gather + Bcast: всі отримують повний результат. MPI_Reduce(sendbuf, recvbuf, count, type, op, root, comm) — редукція з операцією (MPI_SUM, MPI_MAX, MPI_MIN, MPI_PROD). MPI_Allreduce — Reduce + Bcast: всі отримують результат. MPI_Barrier — синхронізація всіх процесів." },
  
        { heading: "MPI Communicators та топології" },
        { text: "MPI дозволяє ділити MPI_COMM_WORLD на підгрупи та створювати власні комунікатори (MPI_Comm_split, MPI_Comm_dup). Топологічні комунікатори: MPI_Cart_create — картезіанська сітка (використовується в алгоритмах Кеннона/Фокса); MPI_Cart_shift — знаходження сусідніх процесів; MPI_Graph_create — довільний граф. Топологічні комунікатори дозволяють бібліотеці оптимізувати фізичне відображення процесів на вузли кластера для мінімізації мережевої відстані." }
      ]
    },
  
    // ----------------------------------------------------------
    {
      id: 40,
      title: "40. Колективні методи обміну повідомленнями стандарту MPI та їх застосування",
      answer: [
        { heading: "Роль колективних операцій" },
        { text: "Колективні операції залучають УСІХ процесів у комунікаторі — жоден не може пропустити. Це гарантує консистентність даних. Ефективні реалізації OpenMPI/MPICH використовують оптимальні алгоритми для кожної операції та топології мережі: двійкове дерево для broadcast (O(log p) кроків); butterfly-алгоритм для AllReduce. Помилка викликати колективну операцію лише у деяких процесах — це призводить до зависання." },
  
        { heading: "MPI_Broadcast" },
        { code: `// Root надсилає свої дані ВСІМ процесам:
  double[] matrixB = (rank == 0) ? loadMatrix() : new double[N*N];
  
  // Всі процеси викликають Bcast:
  MPI.COMM_WORLD.Bcast(matrixB, 0, N * N, MPI.DOUBLE, 0 /*root*/);
  
  // Після Bcast: у КОЖНОГО процесу matrixB містить дані від root
  // Типове застосування: broadcast матриці B у стрічковому алгоритмі` },
  
        { heading: "MPI_Scatter та MPI_Gather" },
        { code: `// Scatter: root ділить рівні частини між усіма
  int rowsPerProc = N / size;
  double[] localA  = new double[rowsPerProc * N];
  
  double[] fullA = (rank == 0) ? loadMatrixA() : null;
  MPI.COMM_WORLD.Scatter(
      fullA,   0, rowsPerProc * N, MPI.DOUBLE, // буфер root
      localA,  0, rowsPerProc * N, MPI.DOUBLE, // буфер кожного
      0 /*root*/
  );
  // Тепер localA[i] містить рядки (rank*rowsPerProc + i) матриці A
  
  // ... обчислення localC ...
  
  // Gather: збираємо результати на root
  double[] fullC = (rank == 0) ? new double[N * N] : null;
  MPI.COMM_WORLD.Gather(
      localC, 0, rowsPerProc * N, MPI.DOUBLE,
      fullC,  0, rowsPerProc * N, MPI.DOUBLE,
      0 /*root*/
  );` },
  
        { heading: "MPI_Reduce та MPI_AllReduce" },
        { code: `// Reduce: обчислити глобальну суму часткових результатів
  double[] localSum  = { computeMyPartialSum() };
  double[] globalSum = new double[1];
  
  // Варіант 1: результат тільки у root:
  MPI.COMM_WORLD.Reduce(localSum,  0, globalSum, 0,
                        1, MPI.DOUBLE, MPI.SUM, 0);
  if (rank == 0) System.out.println("Загальна сума: " + globalSum[0]);
  
  // Варіант 2: результат у ВСІХ процесах:
  MPI.COMM_WORLD.Allreduce(localSum, 0, globalSum, 0,
                           1, MPI.DOUBLE, MPI.SUM);
  
  // Підтримувані операції:
  // MPI.SUM, MPI.PROD, MPI.MAX, MPI.MIN,
  // MPI.LAND (логічне AND), MPI.LOR, MPI.BAND (бітове AND), MPI.BOR` },
  
        { heading: "MPI_AllGather та MPI_AllToAll" },
        { text: "MPI_AllGather — кожен процес надсилає свої дані, і всі отримують повний зібраний масив. Використовується в алгоритмах де кожен процес потребує даних усіх інших (наприклад, пошук глобального максимуму + розповсюдження). MPI_AllToAll — кожен процес надсилає різні дані кожному іншому (персоналізована AllGather). Використовується при матричній транспозиції та FFT." },
        { code: `// AllGather: кожен ділиться своїм рядком зі всіма
  double[] myRow   = computeMyRow();  // N елементів
  double[] allRows = new double[size * N]; // N*p елементів
  
  MPI.COMM_WORLD.Allgather(
      myRow,   0, N, MPI.DOUBLE,
      allRows, 0, N, MPI.DOUBLE
  );
  // allRows[rank*N .. (rank+1)*N-1] = myRow кожного процесу` },
  
        { heading: "MPI_Scatterv та MPI_Gatherv — нерівномірний розподіл" },
        { text: "Scatter та Gather розділяють дані рівно (N/p кожному). Якщо N не ділиться на p або підзадачі різного розміру — використовуємо Scatterv/Gatherv: вектор sendcounts[i] задає скільки даних надіслати/отримати від процесу i; вектор displs[i] задає зміщення в буфері. Типове застосування: нерівномірна декомпозиція (наприклад, перший процес обробляє N mod p зайвих рядків)." }
      ]
    },
  
    // ----------------------------------------------------------
    {
      id: 42,
      title: "42. Розробка ефективних паралельних алгоритмів в OpenMPI",
      answer: [
        { heading: "Ключові принципи ефективності у OpenMPI" },
        { text: "1. Мінімізація комунікацій: кожен виклик MPI-функції має латентність (1–10 мкс для InfiniBand). Краще надіслати одне велике повідомлення, ніж кілька дрібних. 2. Перекриття обчислень та комунікацій: поки дані передаються, обчислювати те, що не залежить від цих даних (неблокуючі операції). 3. Балансування навантаження: всі процеси мають рівну кількість роботи, інакше одні чекають інших у бар'єрах. 4. Ефективна серіалізація даних: MPI_Pack/Unpack або MPI_Type_create_struct для складних структур." },
  
        { heading: "Перекриття обчислень та комунікацій" },
        { code: `// ПОГАНО: послідовно — спочатку recv, потім обчислення
  double[] borderFromNeighbour = new double[N];
  MPI.COMM_WORLD.Recv(borderFromNeighbour, 0, N, MPI.DOUBLE,
      neighbour, 0); // блокуємось
  computeAll(localData, borderFromNeighbour);
  
  // ДОБРЕ: неблокуючий Irecv — одночасно отримуємо і обчислюємо
  double[] borderFromNeighbour = new double[N];
  Request recvReq = MPI.COMM_WORLD.Irecv(
      borderFromNeighbour, 0, N, MPI.DOUBLE, neighbour, 0);
  
  // Поки дані йдуть по мережі — обчислюємо внутрішні елементи
  // (ті, що не залежать від сусіда):
  computeInnerStripe(localData);
  
  // Тепер чекаємо дані від сусіда:
  recvReq.Wait();
  
  // Обчислюємо граничні елементи:
  computeBoundaryStripe(localData, borderFromNeighbour);` },
  
        { heading: "Ефективне вимірювання часу та profile" },
        { code: `// Правильне вимірювання у MPI:
  MPI.COMM_WORLD.Barrier(); // синхронізація перед вимірюванням
  double tStart = MPI.Wtime(); // мікросекундна точність
  
  // ... алгоритм ...
  
  MPI.COMM_WORLD.Barrier();
  double elapsed = MPI.Wtime() - tStart;
  
  // Знаходимо максимальний час (критичний шлях):
  double[] localTime = { elapsed };
  double[] maxTime   = new double[1];
  MPI.COMM_WORLD.Reduce(localTime, 0, maxTime, 0,
      1, MPI.DOUBLE, MPI.MAX, 0);
  if (rank == 0) {
      System.out.printf("Час: %.4f с%n", maxTime[0]);
      System.out.printf("Прискорення: %.2fx%n", seqTime / maxTime[0]);
  }` },
  
        { heading: "Паралельне множення матриць у OpenMPI — повна реалізація" },
        { code: `public class MPIMatMul {
      public static void main(String[] args) throws MPIException {
          MPI.Init(args);
          int rank = MPI.COMM_WORLD.Rank();
          int size = MPI.COMM_WORLD.Size();
          int N = 1000;
          int rowsPerProc = N / size;
  
          double[] A = null, B = new double[N * N], C = null;
          double[] localA = new double[rowsPerProc * N];
          double[] localC = new double[rowsPerProc * N];
  
          if (rank == 0) {
              A = generateMatrix(N);
              C = new double[N * N];
          }
  
          // Розподіл рядків A (по rowsPerProc рядків кожному):
          MPI.COMM_WORLD.Scatter(A, 0, rowsPerProc * N, MPI.DOUBLE,
              localA, 0, rowsPerProc * N, MPI.DOUBLE, 0);
          // Broadcast повної B:
          if (rank != 0) B = new double[N * N];
          else System.arraycopy(generateMatrixB(N), 0, B, 0, N * N);
          MPI.COMM_WORLD.Bcast(B, 0, N * N, MPI.DOUBLE, 0);
  
          // Локальне множення (незалежно від інших):
          MPI.COMM_WORLD.Barrier();
          double t0 = MPI.Wtime();
          for (int i = 0; i < rowsPerProc; i++)
              for (int j = 0; j < N; j++)
                  for (int k = 0; k < N; k++)
                      localC[i * N + j] += localA[i * N + k] * B[k * N + j];
  
          // Збір результатів:
          MPI.COMM_WORLD.Gather(localC, 0, rowsPerProc * N, MPI.DOUBLE,
              C, 0, rowsPerProc * N, MPI.DOUBLE, 0);
  
          double elapsed = MPI.Wtime() - t0;
          double[] maxT = new double[1];
          MPI.COMM_WORLD.Reduce(new double[]{elapsed}, 0, maxT, 0,
              1, MPI.DOUBLE, MPI.MAX, 0);
          if (rank == 0)
              System.out.printf("N=%d, p=%d, час=%.3f с%n",
                  N, size, maxT[0]);
          MPI.Finalize();
      }
  }` }
      ]
    },
  
    // ----------------------------------------------------------
    {
      id: 44,
      title: "44. Паралельна реалізація алгоритму МГУА (метод групового урахування аргументів)",
      answer: [
        { heading: "Що таке МГУА" },
        { text: "МГУА (Метод групового урахування аргументів, англ. GMDH — Group Method of Data Handling) — самоорганізуюча процедура для побудови математичних моделей за даними спостережень. Запропонований Олексієм Івахненком у 1968 р. в Інституті кібернетики НАН України. МГУА — еволюційний алгоритм: будує складні поліноміальні моделі шляхом ітеративного відбору найкращих часткових описів." },
        { text: "Алгоритм МГУА: 1) Вхідні дані: N спостережень з M ознаками. Ділимо на навчальну та перевірочну вибірки. 2) Генерація часткових описів: для кожної пари ознак (xi, xj) будуємо поліном ŷ = a₀ + a₁xi + a₂xj + a₃xi² + a₄xj² + a₅xixj. Методом МНК знаходимо коефіцієнти. 3) Відбір: залишаємо F найкращих описів за критерієм на перевірочній вибірці (MSE, AIC, регулярність). 4) Нові ознаки: виходи відібраних описів стають новими ознаками для наступного шару. 5) Повторюємо до зупинки (погіршення критерію або максимальна глибина)." },
  
        { heading: "Паралелізм у МГУА" },
        { text: "МГУА має природний паралелізм: оцінка кожного часткового опису (побудова та оцінка полінома для пари ознак) є повністю незалежною від інших. На шарі з M ознаками виникає C(M,2) = M(M-1)/2 описів — при M=100 це 4950 незалежних задач. Це embarrassingly parallel задача — ідеальна для паралелізації. Додатковий паралелізм: обчислення матриць (X^T X) та (X^T y) для МНК; пошук F найкращих серед K описів (паралельне сортування)." },
  
        { heading: "Паралельна реалізація з ForkJoinPool (Java)" },
        { code: `public class ParallelGMDH {
  
      // Паралельна оцінка всіх часткових описів одного шару:
      public static List<PartialDescription> evaluateLayer(
              double[][] X, double[] y, int F) {
  
          List<int[]> pairs = generatePairs(X[0].length);
  
          // ForkJoin для паралельної оцінки всіх пар:
          ForkJoinPool pool = ForkJoinPool.commonPool();
          List<PartialDescription> allDescriptions =
              pool.invoke(new EvaluatePairsTask(X, y, pairs, 0, pairs.size()));
  
          // Відбір F найкращих (за MSE на тестовій вибірці):
          allDescriptions.sort(Comparator.comparingDouble(pd -> pd.testMSE));
          return allDescriptions.subList(0, Math.min(F, allDescriptions.size()));
      }
  
      static class EvaluatePairsTask
              extends RecursiveTask<List<PartialDescription>> {
          double[][] X; double[] y;
          List<int[]> pairs;
          int from, to;
          static final int THRESHOLD = 50; // пар на задачу
  
          protected List<PartialDescription> compute() {
              if (to - from <= THRESHOLD) {
                  // Послідовне обчислення свого блоку пар:
                  List<PartialDescription> results = new ArrayList<>();
                  for (int i = from; i < to; i++) {
                      int xi = pairs.get(i)[0], xj = pairs.get(i)[1];
                      results.add(fitPolynomial(X, y, xi, xj));
                  }
                  return results;
              }
              int mid = (from + to) / 2;
              EvaluatePairsTask left  =
                  new EvaluatePairsTask(X, y, pairs, from, mid);
              EvaluatePairsTask right =
                  new EvaluatePairsTask(X, y, pairs, mid, to);
              left.fork();
              List<PartialDescription> rightResult = right.compute();
              List<PartialDescription> leftResult  = left.join();
              leftResult.addAll(rightResult);
              return leftResult;
          }
      }
  
      // МНК для полінома 2-го ступеня від двох ознак:
      static PartialDescription fitPolynomial(
              double[][] X, double[] y, int xi, int xj) {
          // Будуємо матрицю Ф: [1, xi, xj, xi^2, xj^2, xi*xj]
          // Знаходимо a = (Ф^T Ф)^{-1} Ф^T y (МНК)
          // Обчислюємо MSE на тренувальній та тестовій вибірках
          // Повертаємо PartialDescription з коефіцієнтами та MSE
          return new PartialDescription(xi, xj, coefficients, trainMSE, testMSE);
      }
  }` },
  
        { heading: "Паралельна реалізація з MPI (розподілений МГУА)" },
        { text: "При великих наборах даних або великій кількості ознак МГУА можна розподілити між вузлами кластера. Root-процес генерує список всіх пар; за допомогою MPI_Scatter розподіляє пари між процесами; кожен процес оцінює свій блок пар; MPI_Gather збирає результати; Root відбирає F найкращих і переходить до наступного шару. Оскільки МГУА є ітеративним алгоритмом, цей процес повторюється для кожного шару." },
  
        { heading: "Оцінка ефективності паралельного МГУА" },
        { text: "При M ознаках і p потоках: кількість задач = C(M,2) = M(M-1)/2; задачі на потік = M(M-1)/(2p); якщо M >> p — ефективність близька до ідеальної. Вузьке місце: фаза відбору F кращих (потрібна глобальна сортування/пошук). При M=1000 і p=8: ~125,000 пар, 15,625 пар на потік — чудове навантаження. Час оцінки одного шару скорочується в ~8x (мінус накладні на сортування ~O(K log K/p) + gather)." }
      ]
    },
  
    // ----------------------------------------------------------
    {
      id: 46,
      title: "46. Технологія Remote Method Invocation (RMI)",
      answer: [
        { heading: "Концепція RMI" },
        { text: "Java RMI (Remote Method Invocation) — технологія, що дозволяє Java-об'єктам на одній JVM викликати методи об'єктів на іншій JVM, що може знаходитися на іншому комп'ютері. RMI надає абстракцію: виклик методу виглядає майже ідентично локальному, але насправді серіалізує аргументи, передає їх по мережі, десеріалізує на стороні сервера, виконує метод і повертає результат. RMI є частиною стандартної Java SE та вбудована у JDK починаючи з Java 1.1." },
  
        { heading: "Архітектура RMI" },
        { text: "Компоненти RMI: Remote Interface — Java-інтерфейс, що extends Remote; всі методи мають оголошувати throws RemoteException. Server Object — реалізує Remote Interface; extends UnicastRemoteObject або явно exportObject(). RMI Registry — naming service (реєстр); клієнти знаходять сервер за іменем. Stub (клієнтська сторона) — локальний proxy, що серіалізує виклик і надсилає мережею. Skeleton (серверна сторона) — приймає мережевий запит, десеріалізує, викликає реальний метод, серіалізує та надсилає результат. З Java 5 skeleton генерується автоматично; з Java 6 stub також не потребує окремої генерації (rmic)." },
  
        { heading: "Реалізація RMI сервісу для паралельних обчислень" },
        { code: `// 1. Remote Interface:
  public interface MatrixService extends Remote {
      double[][] multiply(double[][] A, double[][] B, int threads)
          throws RemoteException;
      double[] parallelSum(double[][] data) throws RemoteException;
  }
  
  // 2. Server Implementation:
  public class MatrixServiceImpl extends UnicastRemoteObject
          implements MatrixService {
  
      public MatrixServiceImpl() throws RemoteException { super(); }
  
      @Override
      public double[][] multiply(double[][] A, double[][] B, int threads)
              throws RemoteException {
          try {
              return ParallelMatMul.multiply(A, B, threads);
          } catch (Exception e) {
              throw new RemoteException("Помилка обчислення", e);
          }
      }
  
      @Override
      public double[] parallelSum(double[][] data) throws RemoteException {
          return Arrays.stream(data)
              .parallel()
              .mapToDouble(row -> Arrays.stream(row).sum())
              .toArray();
      }
  
      // Запуск сервера:
      public static void main(String[] args) throws Exception {
          Registry registry = LocateRegistry.createRegistry(1099);
          MatrixService service = new MatrixServiceImpl();
          registry.bind("MatrixService", service);
          System.out.println("MatrixService готовий на порту 1099");
      }
  }
  
  // 3. Client:
  public class MatrixClient {
      public static void main(String[] args) throws Exception {
          Registry registry = LocateRegistry.getRegistry("server-host", 1099);
          MatrixService service =
              (MatrixService) registry.lookup("MatrixService");
  
          double[][] A = generateMatrix(500);
          double[][] B = generateMatrix(500);
          double[][] C = service.multiply(A, B, 8); // виконується на сервері!
          System.out.println("Результат отримано: " + C.length + "x" + C[0].length);
      }
  }` },
  
        { heading: "RMI у контексті розподілених паралельних обчислень" },
        { text: "RMI надає простий механізм для розподілення обчислень між вузлами без складного MPI-коду. Master-Worker з RMI: Master реєструє або знаходить Worker-сервіси через Registry; передає підзадачі (блоки матриць) Worker'ам через RMI-виклики; збирає результати. Обмеження: серіалізація Java-об'єктів може бути повільною для великих масивів (порівняно з MPI; буфери пам'яті передаються byte[]); RMI не оптимізований для HPC навантажень; вища затримка ніж MPI (~1–10 мс vs ~1–10 мкс)." },
  
        { heading: "RMI vs сучасні альтернативи" },
        { text: "RMI vs REST (HTTP): REST простіший для крос-мовної взаємодії, але вищий overhead. RMI vs gRPC: gRPC (Protocol Buffers + HTTP/2) — значно ефективніша серіалізація, але складніша у налаштуванні. RMI vs Apache Thrift: Thrift генерує код для багатьох мов, ефективна бінарна серіалізація. У сучасних проектах RMI використовується рідко (замінений gRPC/REST), але залишається актуальним у деяких Java EE компонентах (JMX, Java IDL)." }
      ]
    },
  
    // ----------------------------------------------------------
    {
      id: 48,
      title: "48. Базові складові грід-системи",
      answer: [
        { heading: "Визначення грід-системи" },
        { text: "Грід-система (Grid System) — координована інфраструктура розподілених обчислень, що об'єднує гетерогенні обчислювальні ресурси різних організацій через мережу для спільного вирішення великих задач. Термін аналогічний до електричної мережі (power grid): ресурси надаються 'на вимогу', незалежно від їх фізичного розташування. Ключова концепція — Віртуальна Організація (VO): динамічна група людей та установ, що поділяють ресурси за певними правилами." },
  
        { heading: "Базові складові грід-системи" },
        { text: "Ресурсний рівень (Fabric Layer): фізичні ресурси, що надаються у спільне використання. Обчислювальні ресурси: HPC кластери, робочі станції, хмарні VM. Ресурси зберігання: дискові масиви, стрічкові бібліотеки, розподілені FS (GPFS, Lustre). Мережеві ресурси: канали зв'язку, маршрутизатори, QoS-механізми. Наукові прилади: детектори, телескопи, сенсорні мережі." },
        { text: "Рівень підключення (Connectivity Layer): протоколи для безпечного доступу до ресурсів через адміністративні межі. GSI (Grid Security Infrastructure): X.509 PKI сертифікати; proxy-сертифікати для делегування авторизації; однократна аутентифікація (SSO — Single Sign-On). SSL/TLS: шифрування трафіку. GridFTP: високопродуктивна передача даних (паралельні потоки, відновлення після збою)." },
        { text: "Ресурсний рівень послуг (Resource Layer): управління окремими ресурсами. GRAM (Globus Resource Allocation and Management): надсилання задач на локальний планувальник (PBS, SLURM, SGE); моніторинг стану виконання; управління ресурсами. Інформаційні сервіси (MDS/BDII): реєстрація та виявлення доступних ресурсів (LDAP-based)." },
        { text: "Колективний рівень (Collective Layer): сервіси, що координують кілька ресурсів. Брокер ресурсів (Resource Broker): відбирає оптимальний ресурс для задачі (matchmaking за RSL/JSDL); Планувальник задач: WMS (Workload Management System) у gLite. Каталог даних: LFC (LCG File Catalog) — де знаходяться репліки файлів. Сервіс реплікації: автоматичне копіювання даних між сайтами для ефективного доступу. Моніторинг і облік: збір статистики використання ресурсів." },
        { text: "Рівень застосувань (Application Layer): наукові та інженерні застосування, що використовують грід. Science Gateways: веб-портали для запуску грід-задач без знання технічних деталей. Робочі потоки (Scientific Workflows): Taverna, Pegasus — автоматизація складних багатокрокових обчислень." },
  
        { heading: "Globus Toolkit — де-факто стандарт грід ПЗ" },
        { text: "Globus Toolkit (GT) — відкрите ПЗ для побудови грід-інфраструктури. Ключові компоненти: GRAM5/GT6 — управління завданнями; GridFTP — передача файлів (до 10 Gbps за рахунок паралельних TCP-потоків); MyProxy — делегування сертифікатів; GFTP-Utils — командний рядок для GridFTP; Grid Community Toolkit (GCT) — активно підтримуваний форк після закінчення офіційної підтримки GT." },
  
        { heading: "Приклади реальних грід-систем" },
        { text: "WLCG (Worldwide LHC Computing Grid): 170 обчислювальних центрів у 42 країнах; обробка даних Великого адронного колайдера (ЦЕРН); Tier-0 (CERN) → Tier-1 (11 великих центрів) → Tier-2 (університети). EGI (European Grid Infrastructure): понад 300 центрів у Європі; 1 мільйон ядер; 300 ПБ сховище. OSG (Open Science Grid, США): підтримка науки від астрофізики до біоінформатики. XSEDE/ACCESS (США): інтегрована система національних суперкомп'ютерних центрів." }
      ]
    },
  
    // ----------------------------------------------------------
    {
      id: 50,
      title: "50. Організація і управління розподіленими грід-ресурсами",
      answer: [
        { heading: "Управління ресурсами в грід" },
        { text: "Управління ресурсами в грід — складна задача через гетерогенність (різне апаратне і програмне забезпечення), географічну розподіленість, автономність адміністраторів кожного сайту та динамічність (ресурси можуть бути недоступні без попередження). Ключові завдання: виявлення ресурсів (Resource Discovery); резервування ресурсів (Resource Reservation); планування задач (Job Scheduling); моніторинг та облік (Monitoring & Accounting)." },
  
        { heading: "Виявлення ресурсів (Resource Discovery)" },
        { text: "MDS (Monitoring and Discovery Service) у Globus Toolkit та BDII (Berkeley Database Information Index) у gLite реалізують інформаційну службу ресурсів. GLUE Schema (Grid Laboratory Uniform Environment) — стандартна XML-схема опису ресурсів: ComputingElement (CPU, RAM, OS, завантаженість); StorageElement (ємність, тип, доступні протоколи); NetworkElement (пропускна здатність, затримка). Процес виявлення: агенти на кожному сайті реєструють інформацію в локальний BDII; регіональні BDII агрегують локальні; топ-рівневий BDII має повну картину всієї грід-інфраструктури." },
  
        { heading: "Брокер ресурсів та планування задач" },
        { text: "WMS (Workload Management System) у gLite/EGI та Condor/HTCondor у OSG відповідають за matchmaking — пошук підходящого ресурсу для задачі. Опис задачі в JDL (Job Description Language) або JSDL: вимоги до CPU (тип, кількість, мінімальна частота), пам'яті, ПЗ, VO-членства. WMS порівнює вимоги задачі із доступними ресурсами та обирає оптимальний сайт. Стратегії: FIFO; пріоритетна черга; fair-share (рівномірне використання квоти між VO); backfill (заповнення 'вікон' у розкладі короткими задачами)." },
        { code: `// JDL-опис задачі для WMS:
  [
    Type = "Job";
    JobType = "Normal";
    Executable = "/bin/sh";
    Arguments = "run_simulation.sh";
    StdOutput = "output.txt";
    StdError = "error.txt";
    InputSandbox = {"run_simulation.sh", "input_data.dat"};
    OutputSandbox = {"output.txt", "error.txt", "results.dat"};
    Requirements = (
      Member("VO:physics", other.GlueHostApplicationSoftwareRunTimeEnvironment)
      && other.GlueCEPolicyMaxCPUTime >= 1440
      && other.GlueCEStateFreeCPUs >= 8
    );
    Rank = -other.GlueCEStateFreeCPUs; // вибираємо сайт з max вільними CPU
  ]` },
  
        { heading: "Управління даними в грід" },
        { text: "Управління даними — одна з найскладніших задач в грід. Дані можуть бути терабайтами або петабайтами і зберігатися у кількох копіях (репліках) на різних сайтах. LFC (LCG File Catalog): кожен файл має глобальний ідентифікатор (GUID) та LFN (Logical File Name); реплікаційний каталог відображає LFN → список фізичних місцезнаходжень (SFN — Site File Name). FTS (File Transfer Service): надійний менеджер передачі файлів між сайтами; retry при збоях; пріоритети; квоти. Стратегії реплікації: eager replication (заздалегідь, до виконання задачі); lazy replication (при першому запиті); popularity-based (частіше використовувані дані реплікуються більше)." },
  
        { heading: "Моніторинг та облік" },
        { text: "GridICE, Ganglia, Nagios — моніторинг стану вузлів у реальному часі. R-GMA (Relational Grid Monitoring Architecture) та APEL (Accounting Processor for Event Logs) — грід-рівень моніторингу та збору статистики використання. Апорте (APEL): кожен сайт збирає записи про виконані задачі (CPU-час, пам'ять, користувач, VO); записи надсилаються до центрального APEL сервера для агрегації; звіти допомагають VO планувати ресурси та розраховуватися зі сторонами." },
  
        { heading: "Управління сертифікатами та аутентифікація" },
        { text: "VOMS (Virtual Organization Membership Service): централізований сервіс управління членством у VO; видає VOMS proxy-certificates з атрибутами (VO, роль, group); грід-ресурс перевіряє VOMS-атрибути для авторизації. Workflow: користувач отримує особистий X.509 сертифікат від довіреного CA; voms-proxy-init генерує proxy-cert з VOMS-атрибутами (дійсний 12–24 год); MyProxy зберігає proxy на сервері для довготривалих задач; при виконанні задачі GRAM отримує оновлений proxy з MyProxy без участі користувача." },
  
        { heading: "Грід в контексті Spring Boot застосунку" },
        { text: "Spring Boot застосунок може виступати як грід-клієнт або грід-сервіс: як клієнт — надсилати задачі до грід-брокера через REST API (WMS REST endpoint); отримувати результати після завершення. як сервіс — реєструватися в BDII, приймати задачі через GRAM API, виконувати та повертати результати. Сучасна тенденція: хмарні провайдери (AWS, Azure) пропонують HPC-сервіси (AWS ParallelCluster, Azure CycleCloud), що інтегруються з грід через cloud bursting, дозволяючи Spring Boot мікросервісам надсилати важкі обчислення в 'хмарний грід' при піковому навантаженні." }
      ]
    }
  
  ]; // END questions array
  
  // ============================================================
  // DOCUMENT BUILDER
  // ============================================================
  
  function makeBody(text) {
    return new Paragraph({
      children: [new TextRun({ text, size: 22, font: "Arial" })],
      spacing: { before: 80, after: 80 },
      alignment: AlignmentType.JUSTIFIED
    });
  }
  
  function makeSubheading(text) {
    return new Paragraph({
      children: [new TextRun({ text, bold: true, size: 24, font: "Arial", color: "1F4E79" })],
      spacing: { before: 180, after: 60 }
    });
  }
  
  function makeCodeLine(line) {
    return new Paragraph({
      children: [new TextRun({ text: line, size: 17, font: "Courier New", color: "1A1A2E" })],
      spacing: { before: 0, after: 0 },
      indent: { left: 560 }
    });
  }
  
  function makeDivider() {
    return new Paragraph({
      children: [],
      border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: "CCCCCC", space: 1 } },
      spacing: { before: 240, after: 240 }
    });
  }
  
  const children = [];
  
  // Title page
  children.push(
    new Paragraph({
      children: [new TextRun({ text: "Технологія паралельних обчислень", bold: true, size: 48, font: "Arial", color: "1F4E79" })],
      alignment: AlignmentType.CENTER,
      spacing: { before: 2000, after: 400 }
    }),
    new Paragraph({
      children: [new TextRun({ text: "Відповіді на екзаменаційні питання", size: 32, font: "Arial", color: "2E75B6" })],
      alignment: AlignmentType.CENTER,
      spacing: { before: 200, after: 200 }
    }),
    new Paragraph({
      children: [new TextRun({ text: "Парні питання (2, 4, 6 ... 50)", size: 26, font: "Arial", color: "555555" })],
      alignment: AlignmentType.CENTER,
      spacing: { before: 100, after: 200 }
    }),
    new Paragraph({
      children: [new TextRun({ text: "Мова програмування: Java", size: 24, font: "Arial", color: "777777" })],
      alignment: AlignmentType.CENTER,
      spacing: { before: 100, after: 4000 }
    })
  );
  
  // Generate all question sections
  for (let qi = 0; qi < questions.length; qi++) {
    const q = questions[qi];
  
    if (qi > 0) {
      children.push(new Paragraph({ children: [new PageBreak()] }));
    }
  
    // Question heading
    children.push(
      new Paragraph({
        heading: HeadingLevel.HEADING_1,
        children: [new TextRun({ text: q.title, bold: true, size: 30, font: "Arial", color: "1F4E79" })],
        spacing: { before: 300, after: 200 }
      })
    );
  
    // Answer blocks
    for (const block of q.answer) {
      if (block.heading) {
        children.push(makeSubheading(block.heading));
      } else if (block.text) {
        children.push(makeBody(block.text));
      } else if (block.code) {
        const lines = block.code.split('\n');
        for (const line of lines) {
          children.push(makeCodeLine(line));
        }
        children.push(new Paragraph({ children: [], spacing: { before: 60, after: 60 } }));
      }
    }
  
    children.push(makeDivider());
  }
  
  // ============================================================
  // BUILD DOCUMENT
  // ============================================================
  const doc = new Document({
    styles: {
      default: {
        document: { run: { font: "Arial", size: 22 } }
      },
      paragraphStyles: [
        {
          id: "Heading1",
          name: "Heading 1",
          basedOn: "Normal",
          next: "Normal",
          quickFormat: true,
          run: { size: 30, bold: true, font: "Arial", color: "1F4E79" },
          paragraph: { spacing: { before: 300, after: 200 }, outlineLevel: 0 }
        }
      ]
    },
    sections: [{
      properties: {
        page: {
          size: { width: 11906, height: 16838 }, // A4
          margin: { top: 1134, right: 1134, bottom: 1134, left: 1134 }
        }
      },
      children: children
    }]
  });
  
  // ============================================================
  // SAVE FILE
  // ============================================================
  const OUTPUT_PATH = './docs/ev1.docx';
  
  Packer.toBuffer(doc).then(buffer => {
    fs.writeFileSync(OUTPUT_PATH, buffer);
  }).catch(err => {
    console.error('Error generating document:', err);
    process.exit(1);
  });