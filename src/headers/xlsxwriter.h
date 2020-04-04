# 1 "/usr/local/include/xlsxwriter.h"
# 1 "<built-in>"
# 1 "<command-line>"
# 31 "<command-line>"
# 1 "/usr/include/stdc-predef.h" 1 3 4
# 32 "<command-line>" 2
# 1 "/usr/local/include/xlsxwriter.h"
# 16 "/usr/local/include/xlsxwriter.h"
# 1 "/usr/local/include/xlsxwriter/workbook.h" 1
# 44 "/usr/local/include/xlsxwriter/workbook.h"
# 1 "/usr/include/stdint.h" 1 3 4
# 20 "/usr/include/stdint.h" 3 4
# 1 "/usr/include/bits/alltypes.h" 1 3 4
# 106 "/usr/include/bits/alltypes.h" 3 4

# 106 "/usr/include/bits/alltypes.h" 3 4
typedef unsigned long uintptr_t;
# 121 "/usr/include/bits/alltypes.h" 3 4
typedef long intptr_t;
# 137 "/usr/include/bits/alltypes.h" 3 4
typedef signed char int8_t;




typedef signed short int16_t;




typedef signed int int32_t;




typedef signed long int64_t;




typedef signed long intmax_t;




typedef unsigned char uint8_t;




typedef unsigned short uint16_t;




typedef unsigned int uint32_t;




typedef unsigned long uint64_t;
# 187 "/usr/include/bits/alltypes.h" 3 4
typedef unsigned long uintmax_t;
# 21 "/usr/include/stdint.h" 2 3 4

typedef int8_t int_fast8_t;
typedef int64_t int_fast64_t;

typedef int8_t int_least8_t;
typedef int16_t int_least16_t;
typedef int32_t int_least32_t;
typedef int64_t int_least64_t;

typedef uint8_t uint_fast8_t;
typedef uint64_t uint_fast64_t;

typedef uint8_t uint_least8_t;
typedef uint16_t uint_least16_t;
typedef uint32_t uint_least32_t;
typedef uint64_t uint_least64_t;
# 95 "/usr/include/stdint.h" 3 4
# 1 "/usr/include/bits/stdint.h" 1 3 4
typedef int32_t int_fast16_t;
typedef int32_t int_fast32_t;
typedef uint32_t uint_fast16_t;
typedef uint32_t uint_fast32_t;
# 96 "/usr/include/stdint.h" 2 3 4
# 45 "/usr/local/include/xlsxwriter/workbook.h" 2
# 1 "/usr/include/stdio.h" 1 3 4







# 1 "/usr/include/features.h" 1 3 4
# 9 "/usr/include/stdio.h" 2 3 4
# 26 "/usr/include/stdio.h" 3 4
# 1 "/usr/include/bits/alltypes.h" 1 3 4





typedef __builtin_va_list va_list;




typedef __builtin_va_list __isoc_va_list;
# 101 "/usr/include/bits/alltypes.h" 3 4
typedef unsigned long size_t;
# 116 "/usr/include/bits/alltypes.h" 3 4
typedef long ssize_t;
# 203 "/usr/include/bits/alltypes.h" 3 4
typedef long off_t;
# 361 "/usr/include/bits/alltypes.h" 3 4
typedef struct _IO_FILE FILE;
# 27 "/usr/include/stdio.h" 2 3 4
# 54 "/usr/include/stdio.h" 3 4
typedef union _G_fpos64_t {
 char __opaque[16];
 long long __lldata;
 double __align;
} fpos_t;

extern FILE *const stdin;
extern FILE *const stdout;
extern FILE *const stderr;





FILE *fopen(const char *restrict, const char *restrict);
FILE *freopen(const char *restrict, const char *restrict, FILE *restrict);
int fclose(FILE *);

int remove(const char *);
int rename(const char *, const char *);

int feof(FILE *);
int ferror(FILE *);
int fflush(FILE *);
void clearerr(FILE *);

int fseek(FILE *, long, int);
long ftell(FILE *);
void rewind(FILE *);

int fgetpos(FILE *restrict, fpos_t *restrict);
int fsetpos(FILE *, const fpos_t *);

size_t fread(void *restrict, size_t, size_t, FILE *restrict);
size_t fwrite(const void *restrict, size_t, size_t, FILE *restrict);

int fgetc(FILE *);
int getc(FILE *);
int getchar(void);
int ungetc(int, FILE *);

int fputc(int, FILE *);
int putc(int, FILE *);
int putchar(int);

char *fgets(char *restrict, int, FILE *restrict);




int fputs(const char *restrict, FILE *restrict);
int puts(const char *);

int printf(const char *restrict, ...);
int fprintf(FILE *restrict, const char *restrict, ...);
int sprintf(char *restrict, const char *restrict, ...);
int snprintf(char *restrict, size_t, const char *restrict, ...);

int vprintf(const char *restrict, __isoc_va_list);
int vfprintf(FILE *restrict, const char *restrict, __isoc_va_list);
int vsprintf(char *restrict, const char *restrict, __isoc_va_list);
int vsnprintf(char *restrict, size_t, const char *restrict, __isoc_va_list);

int scanf(const char *restrict, ...);
int fscanf(FILE *restrict, const char *restrict, ...);
int sscanf(const char *restrict, const char *restrict, ...);
int vscanf(const char *restrict, __isoc_va_list);
int vfscanf(FILE *restrict, const char *restrict, __isoc_va_list);
int vsscanf(const char *restrict, const char *restrict, __isoc_va_list);

void perror(const char *);

int setvbuf(FILE *restrict, char *restrict, int, size_t);
void setbuf(FILE *restrict, char *restrict);

char *tmpnam(char *);
FILE *tmpfile(void);




FILE *fmemopen(void *restrict, size_t, const char *restrict);
FILE *open_memstream(char **, size_t *);
FILE *fdopen(int, const char *);
FILE *popen(const char *, const char *);
int pclose(FILE *);
int fileno(FILE *);
int fseeko(FILE *, off_t, int);
off_t ftello(FILE *);
int dprintf(int, const char *restrict, ...);
int vdprintf(int, const char *restrict, __isoc_va_list);
void flockfile(FILE *);
int ftrylockfile(FILE *);
void funlockfile(FILE *);
int getc_unlocked(FILE *);
int getchar_unlocked(void);
int putc_unlocked(int, FILE *);
int putchar_unlocked(int);
ssize_t getdelim(char **restrict, size_t *restrict, int, FILE *restrict);
ssize_t getline(char **restrict, size_t *restrict, FILE *restrict);
int renameat(int, const char *, int, const char *);
char *ctermid(char *);







char *tempnam(const char *, const char *);




char *cuserid(char *);
void setlinebuf(FILE *);
void setbuffer(FILE *, char *, size_t);
int fgetc_unlocked(FILE *);
int fputc_unlocked(int, FILE *);
int fflush_unlocked(FILE *);
size_t fread_unlocked(void *, size_t, size_t, FILE *);
size_t fwrite_unlocked(const void *, size_t, size_t, FILE *);
void clearerr_unlocked(FILE *);
int feof_unlocked(FILE *);
int ferror_unlocked(FILE *);
int fileno_unlocked(FILE *);
int getw(FILE *);
int putw(int, FILE *);
char *fgetln(FILE *, size_t *);
int asprintf(char **, const char *, ...);
int vasprintf(char **, const char *, __isoc_va_list);
# 46 "/usr/local/include/xlsxwriter/workbook.h" 2
# 1 "/usr/include/errno.h" 1 3 4
# 10 "/usr/include/errno.h" 3 4
# 1 "/usr/include/bits/errno.h" 1 3 4
# 11 "/usr/include/errno.h" 2 3 4


__attribute__((const))

int *__errno_location(void);
# 47 "/usr/local/include/xlsxwriter/workbook.h" 2

# 1 "/usr/local/include/xlsxwriter/worksheet.h" 1
# 47 "/usr/local/include/xlsxwriter/worksheet.h"
# 1 "/usr/include/stdlib.h" 1 3 4
# 19 "/usr/include/stdlib.h" 3 4
# 1 "/usr/include/bits/alltypes.h" 1 3 4
# 18 "/usr/include/bits/alltypes.h" 3 4
typedef int wchar_t;
# 20 "/usr/include/stdlib.h" 2 3 4

int atoi (const char *);
long atol (const char *);
long long atoll (const char *);
double atof (const char *);

float strtof (const char *restrict, char **restrict);
double strtod (const char *restrict, char **restrict);
long double strtold (const char *restrict, char **restrict);

long strtol (const char *restrict, char **restrict, int);
unsigned long strtoul (const char *restrict, char **restrict, int);
long long strtoll (const char *restrict, char **restrict, int);
unsigned long long strtoull (const char *restrict, char **restrict, int);

int rand (void);
void srand (unsigned);

void *malloc (size_t);
void *calloc (size_t, size_t);
void *realloc (void *, size_t);
void free (void *);
void *aligned_alloc(size_t, size_t);

_Noreturn void abort (void);
int atexit (void (*) (void));
_Noreturn void exit (int);
_Noreturn void _Exit (int);
int at_quick_exit (void (*) (void));
_Noreturn void quick_exit (int);

char *getenv (const char *);

int system (const char *);

void *bsearch (const void *, const void *, size_t, size_t, int (*)(const void *, const void *));
void qsort (void *, size_t, size_t, int (*)(const void *, const void *));

int abs (int);
long labs (long);
long long llabs (long long);

typedef struct { int quot, rem; } div_t;
typedef struct { long quot, rem; } ldiv_t;
typedef struct { long long quot, rem; } lldiv_t;

div_t div (int, int);
ldiv_t ldiv (long, long);
lldiv_t lldiv (long long, long long);

int mblen (const char *, size_t);
int mbtowc (wchar_t *restrict, const char *restrict, size_t);
int wctomb (char *, wchar_t);
size_t mbstowcs (wchar_t *restrict, const char *restrict, size_t);
size_t wcstombs (char *restrict, const wchar_t *restrict, size_t);




size_t __ctype_get_mb_cur_max(void);
# 99 "/usr/include/stdlib.h" 3 4
int posix_memalign (void **, size_t, size_t);
int setenv (const char *, const char *, int);
int unsetenv (const char *);
int mkstemp (char *);
int mkostemp (char *, int);
char *mkdtemp (char *);
int getsubopt (char **, char *const *, char **);
int rand_r (unsigned *);






char *realpath (const char *restrict, char *restrict);
long int random (void);
void srandom (unsigned int);
char *initstate (unsigned int, char *, size_t);
char *setstate (char *);
int putenv (char *);
int posix_openpt (int);
int grantpt (int);
int unlockpt (int);
char *ptsname (int);
char *l64a (long);
long a64l (const char *);
void setkey (const char *);
double drand48 (void);
double erand48 (unsigned short [3]);
long int lrand48 (void);
long int nrand48 (unsigned short [3]);
long mrand48 (void);
long jrand48 (unsigned short [3]);
void srand48 (long);
unsigned short *seed48 (unsigned short [3]);
void lcong48 (unsigned short [7]);



# 1 "/app/src/headers/override/alloca.h" 1 3 4
# 139 "/usr/include/stdlib.h" 2 3 4
char *mktemp (char *);
int mkstemps (char *, int);
int mkostemps (char *, int, int);
void *valloc (size_t);
void *memalign(size_t, size_t);
int getloadavg(double *, int);
int clearenv(void);
# 48 "/usr/local/include/xlsxwriter/worksheet.h" 2

# 1 "/usr/include/string.h" 1 3 4
# 23 "/usr/include/string.h" 3 4
# 1 "/usr/include/bits/alltypes.h" 1 3 4
# 373 "/usr/include/bits/alltypes.h" 3 4
typedef struct __locale_struct * locale_t;
# 24 "/usr/include/string.h" 2 3 4

void *memcpy (void *restrict, const void *restrict, size_t);
void *memmove (void *, const void *, size_t);
void *memset (void *, int, size_t);
int memcmp (const void *, const void *, size_t);
void *memchr (const void *, int, size_t);

char *strcpy (char *restrict, const char *restrict);
char *strncpy (char *restrict, const char *restrict, size_t);

char *strcat (char *restrict, const char *restrict);
char *strncat (char *restrict, const char *restrict, size_t);

int strcmp (const char *, const char *);
int strncmp (const char *, const char *, size_t);

int strcoll (const char *, const char *);
size_t strxfrm (char *restrict, const char *restrict, size_t);

char *strchr (const char *, int);
char *strrchr (const char *, int);

size_t strcspn (const char *, const char *);
size_t strspn (const char *, const char *);
char *strpbrk (const char *, const char *);
char *strstr (const char *, const char *);
char *strtok (char *restrict, const char *restrict);

size_t strlen (const char *);

char *strerror (int);


# 1 "/usr/include/strings.h" 1 3 4
# 11 "/usr/include/strings.h" 3 4
# 1 "/usr/include/bits/alltypes.h" 1 3 4
# 12 "/usr/include/strings.h" 2 3 4




int bcmp (const void *, const void *, size_t);
void bcopy (const void *, void *, size_t);
void bzero (void *, size_t);
char *index (const char *, int);
char *rindex (const char *, int);



int ffs (int);
int ffsl (long);
int ffsll (long long);


int strcasecmp (const char *, const char *);
int strncasecmp (const char *, const char *, size_t);

int strcasecmp_l (const char *, const char *, locale_t);
int strncasecmp_l (const char *, const char *, size_t, locale_t);
# 58 "/usr/include/string.h" 2 3 4





char *strtok_r (char *restrict, const char *restrict, char **restrict);
int strerror_r (int, char *, size_t);
char *stpcpy(char *restrict, const char *restrict);
char *stpncpy(char *restrict, const char *restrict, size_t);
size_t strnlen (const char *, size_t);
char *strdup (const char *);
char *strndup (const char *, size_t);
char *strsignal(int);
char *strerror_l (int, locale_t);
int strcoll_l (const char *, const char *, locale_t);
size_t strxfrm_l (char *restrict, const char *restrict, size_t, locale_t);




void *memccpy (void *restrict, const void *restrict, int, size_t);



char *strsep(char **, const char *);
size_t strlcat (char *, const char *, size_t);
size_t strlcpy (char *, const char *, size_t);
void explicit_bzero (void *, size_t);
# 50 "/usr/local/include/xlsxwriter/worksheet.h" 2

# 1 "/usr/local/include/xlsxwriter/shared_strings.h" 1
# 16 "/usr/local/include/xlsxwriter/shared_strings.h"
# 1 "/usr/local/include/xlsxwriter/common.h" 1
# 18 "/usr/local/include/xlsxwriter/common.h"
# 1 "/usr/include/time.h" 1 3 4
# 31 "/usr/include/time.h" 3 4
# 1 "/usr/include/bits/alltypes.h" 1 3 4
# 55 "/usr/include/bits/alltypes.h" 3 4
typedef long time_t;
# 250 "/usr/include/bits/alltypes.h" 3 4
typedef void * timer_t;




typedef int clockid_t;




typedef long clock_t;
# 270 "/usr/include/bits/alltypes.h" 3 4
struct timespec { time_t tv_sec; long tv_nsec; };





typedef int pid_t;
# 32 "/usr/include/time.h" 2 3 4






struct tm {
 int tm_sec;
 int tm_min;
 int tm_hour;
 int tm_mday;
 int tm_mon;
 int tm_year;
 int tm_wday;
 int tm_yday;
 int tm_isdst;
 long tm_gmtoff;
 const char *tm_zone;
};

clock_t clock (void);
time_t time (time_t *);
double difftime (time_t, time_t);
time_t mktime (struct tm *);
size_t strftime (char *restrict, size_t, const char *restrict, const struct tm *restrict);
struct tm *gmtime (const time_t *);
struct tm *localtime (const time_t *);
char *asctime (const struct tm *);
char *ctime (const time_t *);
int timespec_get(struct timespec *, int);
# 71 "/usr/include/time.h" 3 4
size_t strftime_l (char * restrict, size_t, const char * restrict, const struct tm * restrict, locale_t);

struct tm *gmtime_r (const time_t *restrict, struct tm *restrict);
struct tm *localtime_r (const time_t *restrict, struct tm *restrict);
char *asctime_r (const struct tm *restrict, char *restrict);
char *ctime_r (const time_t *, char *);

void tzset (void);

struct itimerspec {
 struct timespec it_interval;
 struct timespec it_value;
};
# 100 "/usr/include/time.h" 3 4
int nanosleep (const struct timespec *, struct timespec *);
int clock_getres (clockid_t, struct timespec *);
int clock_gettime (clockid_t, struct timespec *);
int clock_settime (clockid_t, const struct timespec *);
int clock_nanosleep (clockid_t, int, const struct timespec *, struct timespec *);
int clock_getcpuclockid (pid_t, clockid_t *);

struct sigevent;
int timer_create (clockid_t, struct sigevent *restrict, timer_t *restrict);
int timer_delete (timer_t);
int timer_settime (timer_t, int, const struct itimerspec *restrict, struct itimerspec *restrict);
int timer_gettime (timer_t, struct itimerspec *);
int timer_getoverrun (timer_t);

extern char *tzname[2];





char *strptime (const char *restrict, const char *restrict, struct tm *restrict);
extern int daylight;
extern long timezone;
extern int getdate_err;
struct tm *getdate (const char *);




int stime(const time_t *);
time_t timegm(struct tm *);
# 19 "/usr/local/include/xlsxwriter/common.h" 2
# 1 "/usr/local/include/xlsxwriter/third_party/queue.h" 1
# 20 "/usr/local/include/xlsxwriter/common.h" 2
# 1 "/usr/local/include/xlsxwriter/third_party/tree.h" 1
# 21 "/usr/local/include/xlsxwriter/common.h" 2
# 40 "/usr/local/include/xlsxwriter/common.h"

# 40 "/usr/local/include/xlsxwriter/common.h"
typedef uint32_t lxw_row_t;





typedef uint16_t lxw_col_t;


enum lxw_boolean {

    LXW_FALSE,

    LXW_TRUE
};







typedef enum lxw_error {


    LXW_NO_ERROR = 0,


    LXW_ERROR_MEMORY_MALLOC_FAILED,


    LXW_ERROR_CREATING_XLSX_FILE,


    LXW_ERROR_CREATING_TMPFILE,


    LXW_ERROR_READING_TMPFILE,


    LXW_ERROR_ZIP_FILE_OPERATION,


    LXW_ERROR_ZIP_PARAMETER_ERROR,


    LXW_ERROR_ZIP_BAD_ZIP_FILE,


    LXW_ERROR_ZIP_INTERNAL_ERROR,


    LXW_ERROR_ZIP_FILE_ADD,


    LXW_ERROR_ZIP_CLOSE,


    LXW_ERROR_FEATURE_NOT_SUPPORTED,


    LXW_ERROR_NULL_PARAMETER_IGNORED,


    LXW_ERROR_PARAMETER_VALIDATION,


    LXW_ERROR_SHEETNAME_LENGTH_EXCEEDED,


    LXW_ERROR_INVALID_SHEETNAME_CHARACTER,


    LXW_ERROR_SHEETNAME_START_END_APOSTROPHE,


    LXW_ERROR_SHEETNAME_ALREADY_USED,


    LXW_ERROR_SHEETNAME_RESERVED,


    LXW_ERROR_32_STRING_LENGTH_EXCEEDED,


    LXW_ERROR_128_STRING_LENGTH_EXCEEDED,


    LXW_ERROR_255_STRING_LENGTH_EXCEEDED,


    LXW_ERROR_MAX_STRING_LENGTH_EXCEEDED,


    LXW_ERROR_SHARED_STRING_INDEX_NOT_FOUND,


    LXW_ERROR_WORKSHEET_INDEX_OUT_OF_RANGE,


    LXW_ERROR_WORKSHEET_MAX_URL_LENGTH_EXCEEDED,


    LXW_ERROR_WORKSHEET_MAX_NUMBER_URLS_EXCEEDED,


    LXW_ERROR_IMAGE_DIMENSIONS,

    LXW_MAX_ERRNO
} lxw_error;





typedef struct lxw_datetime {


    int year;

    int month;

    int day;

    int hour;

    int min;

    double sec;

} lxw_datetime;

enum lxw_custom_property_types {
    LXW_CUSTOM_NONE,
    LXW_CUSTOM_STRING,
    LXW_CUSTOM_DOUBLE,
    LXW_CUSTOM_INTEGER,
    LXW_CUSTOM_BOOLEAN,
    LXW_CUSTOM_DATETIME
};
# 376 "/usr/local/include/xlsxwriter/common.h"
struct lxw_formats { struct lxw_format *stqh_first; struct lxw_format **stqh_last; };


struct lxw_tuples { struct lxw_tuple *stqh_first; struct lxw_tuple **stqh_last; };
struct lxw_custom_properties { struct lxw_custom_property *stqh_first; struct lxw_custom_property **stqh_last; };

typedef struct lxw_tuple {
    char *key;
    char *value;

    struct { struct lxw_tuple *stqe_next; } list_pointers;
} lxw_tuple;


typedef struct lxw_custom_property {

    enum lxw_custom_property_types type;
    char *name;

    union {
        char *string;
        double number;
        int32_t integer;
        uint8_t boolean;
        lxw_datetime datetime;
    } u;

    struct { struct lxw_custom_property *stqe_next; } list_pointers;

} lxw_custom_property;
# 17 "/usr/local/include/xlsxwriter/shared_strings.h" 2


struct sst_rb_tree { struct sst_element *rbh_root; };


struct sst_order_list { struct sst_element *stqh_first; struct sst_element **stqh_last; };
# 37 "/usr/local/include/xlsxwriter/shared_strings.h"
struct sst_element {
    uint32_t index;
    char *string;
    uint8_t is_rich_string;

    struct { struct sst_element *stqe_next; } sst_order_pointers;
    struct { struct sst_element *rbe_left; struct sst_element *rbe_right; struct sst_element *rbe_parent; int rbe_color; } sst_tree_pointers;
};




typedef struct lxw_sst {
    FILE *file;

    uint32_t string_count;
    uint32_t unique_count;

    struct sst_order_list *order_list;
    struct sst_rb_tree *rb_tree;

} lxw_sst;







lxw_sst *lxw_sst_new(void);
void lxw_sst_free(lxw_sst *sst);
struct sst_element *lxw_get_sst_index(lxw_sst *sst, const char *string,
                                      uint8_t is_rich_string);
void lxw_sst_assemble_xml_file(lxw_sst *self);
# 52 "/usr/local/include/xlsxwriter/worksheet.h" 2
# 1 "/usr/local/include/xlsxwriter/chart.h" 1
# 78 "/usr/local/include/xlsxwriter/chart.h"
# 1 "/usr/local/include/xlsxwriter/format.h" 1
# 65 "/usr/local/include/xlsxwriter/format.h"
# 1 "/usr/local/include/xlsxwriter/hash_table.h" 1
# 20 "/usr/local/include/xlsxwriter/hash_table.h"
struct lxw_hash_order_list { struct lxw_hash_element *stqh_first; struct lxw_hash_element **stqh_last; };
struct lxw_hash_bucket_list { struct lxw_hash_element *slh_first; };


typedef struct lxw_hash_table {
    uint32_t num_buckets;
    uint32_t used_buckets;
    uint32_t unique_count;
    uint8_t free_key;
    uint8_t free_value;

    struct lxw_hash_order_list *order_list;
    struct lxw_hash_bucket_list **buckets;
} lxw_hash_table;
# 42 "/usr/local/include/xlsxwriter/hash_table.h"
typedef struct lxw_hash_element {
    void *key;
    void *value;

    struct { struct lxw_hash_element *stqe_next; } lxw_hash_order_pointers;
    struct { struct lxw_hash_element *sle_next; } lxw_hash_list_pointers;
} lxw_hash_element;
# 57 "/usr/local/include/xlsxwriter/hash_table.h"
lxw_hash_element *lxw_hash_key_exists(lxw_hash_table *lxw_hash, void *key,
                                      size_t key_len);
lxw_hash_element *lxw_insert_hash_element(lxw_hash_table *lxw_hash, void *key,
                                          void *value, size_t key_len);
lxw_hash_table *lxw_hash_new(uint32_t num_buckets, uint8_t free_key,
                             uint8_t free_value);
void lxw_hash_free(lxw_hash_table *lxw_hash);
# 66 "/usr/local/include/xlsxwriter/format.h" 2
# 75 "/usr/local/include/xlsxwriter/format.h"
typedef uint32_t lxw_color_t;
# 94 "/usr/local/include/xlsxwriter/format.h"
enum lxw_format_underlines {
    LXW_UNDERLINE_NONE = 0,


    LXW_UNDERLINE_SINGLE,


    LXW_UNDERLINE_DOUBLE,


    LXW_UNDERLINE_SINGLE_ACCOUNTING,


    LXW_UNDERLINE_DOUBLE_ACCOUNTING
};


enum lxw_format_scripts {


    LXW_FONT_SUPERSCRIPT = 1,


    LXW_FONT_SUBSCRIPT
};


enum lxw_format_alignments {

    LXW_ALIGN_NONE = 0,


    LXW_ALIGN_LEFT,


    LXW_ALIGN_CENTER,


    LXW_ALIGN_RIGHT,


    LXW_ALIGN_FILL,


    LXW_ALIGN_JUSTIFY,


    LXW_ALIGN_CENTER_ACROSS,


    LXW_ALIGN_DISTRIBUTED,


    LXW_ALIGN_VERTICAL_TOP,


    LXW_ALIGN_VERTICAL_BOTTOM,


    LXW_ALIGN_VERTICAL_CENTER,


    LXW_ALIGN_VERTICAL_JUSTIFY,


    LXW_ALIGN_VERTICAL_DISTRIBUTED
};

enum lxw_format_diagonal_types {
    LXW_DIAGONAL_BORDER_UP = 1,
    LXW_DIAGONAL_BORDER_DOWN,
    LXW_DIAGONAL_BORDER_UP_DOWN
};


enum lxw_defined_colors {

    LXW_COLOR_BLACK = 0x1000000,


    LXW_COLOR_BLUE = 0x0000FF,


    LXW_COLOR_BROWN = 0x800000,


    LXW_COLOR_CYAN = 0x00FFFF,


    LXW_COLOR_GRAY = 0x808080,


    LXW_COLOR_GREEN = 0x008000,


    LXW_COLOR_LIME = 0x00FF00,


    LXW_COLOR_MAGENTA = 0xFF00FF,


    LXW_COLOR_NAVY = 0x000080,


    LXW_COLOR_ORANGE = 0xFF6600,


    LXW_COLOR_PINK = 0xFF00FF,


    LXW_COLOR_PURPLE = 0x800080,


    LXW_COLOR_RED = 0xFF0000,


    LXW_COLOR_SILVER = 0xC0C0C0,


    LXW_COLOR_WHITE = 0xFFFFFF,


    LXW_COLOR_YELLOW = 0xFFFF00
};


enum lxw_format_patterns {

    LXW_PATTERN_NONE = 0,


    LXW_PATTERN_SOLID,


    LXW_PATTERN_MEDIUM_GRAY,


    LXW_PATTERN_DARK_GRAY,


    LXW_PATTERN_LIGHT_GRAY,


    LXW_PATTERN_DARK_HORIZONTAL,


    LXW_PATTERN_DARK_VERTICAL,


    LXW_PATTERN_DARK_DOWN,


    LXW_PATTERN_DARK_UP,


    LXW_PATTERN_DARK_GRID,


    LXW_PATTERN_DARK_TRELLIS,


    LXW_PATTERN_LIGHT_HORIZONTAL,


    LXW_PATTERN_LIGHT_VERTICAL,


    LXW_PATTERN_LIGHT_DOWN,


    LXW_PATTERN_LIGHT_UP,


    LXW_PATTERN_LIGHT_GRID,


    LXW_PATTERN_LIGHT_TRELLIS,


    LXW_PATTERN_GRAY_125,


    LXW_PATTERN_GRAY_0625
};


enum lxw_format_borders {

    LXW_BORDER_NONE,


    LXW_BORDER_THIN,


    LXW_BORDER_MEDIUM,


    LXW_BORDER_DASHED,


    LXW_BORDER_DOTTED,


    LXW_BORDER_THICK,


    LXW_BORDER_DOUBLE,


    LXW_BORDER_HAIR,


    LXW_BORDER_MEDIUM_DASHED,


    LXW_BORDER_DASH_DOT,


    LXW_BORDER_MEDIUM_DASH_DOT,


    LXW_BORDER_DASH_DOT_DOT,


    LXW_BORDER_MEDIUM_DASH_DOT_DOT,


    LXW_BORDER_SLANT_DASH_DOT
};
# 348 "/usr/local/include/xlsxwriter/format.h"
typedef struct lxw_format {

    FILE *file;

    lxw_hash_table *xf_format_indices;
    uint16_t *num_xf_formats;

    int32_t xf_index;
    int32_t dxf_index;
    int32_t xf_id;

    char num_format[128];
    char font_name[128];
    char font_scheme[128];
    uint16_t num_format_index;
    uint16_t font_index;
    uint8_t has_font;
    uint8_t has_dxf_font;
    double font_size;
    uint8_t bold;
    uint8_t italic;
    lxw_color_t font_color;
    uint8_t underline;
    uint8_t font_strikeout;
    uint8_t font_outline;
    uint8_t font_shadow;
    uint8_t font_script;
    uint8_t font_family;
    uint8_t font_charset;
    uint8_t font_condense;
    uint8_t font_extend;
    uint8_t theme;
    uint8_t hyperlink;

    uint8_t hidden;
    uint8_t locked;

    uint8_t text_h_align;
    uint8_t text_wrap;
    uint8_t text_v_align;
    uint8_t text_justlast;
    int16_t rotation;

    lxw_color_t fg_color;
    lxw_color_t bg_color;
    uint8_t pattern;
    uint8_t has_fill;
    uint8_t has_dxf_fill;
    int32_t fill_index;
    int32_t fill_count;

    int32_t border_index;
    uint8_t has_border;
    uint8_t has_dxf_border;
    int32_t border_count;

    uint8_t bottom;
    uint8_t diag_border;
    uint8_t diag_type;
    uint8_t left;
    uint8_t right;
    uint8_t top;
    lxw_color_t bottom_color;
    lxw_color_t diag_color;
    lxw_color_t left_color;
    lxw_color_t right_color;
    lxw_color_t top_color;

    uint8_t indent;
    uint8_t shrink;
    uint8_t merge_range;
    uint8_t reading_order;
    uint8_t just_distrib;
    uint8_t color_indexed;
    uint8_t font_only;

    struct { struct lxw_format *stqe_next; } list_pointers;
} lxw_format;




typedef struct lxw_font {

    char font_name[128];
    double font_size;
    uint8_t bold;
    uint8_t italic;
    uint8_t underline;
    uint8_t theme;
    uint8_t font_strikeout;
    uint8_t font_outline;
    uint8_t font_shadow;
    uint8_t font_script;
    uint8_t font_family;
    uint8_t font_charset;
    uint8_t font_condense;
    uint8_t font_extend;
    lxw_color_t font_color;
} lxw_font;




typedef struct lxw_border {

    uint8_t bottom;
    uint8_t diag_border;
    uint8_t diag_type;
    uint8_t left;
    uint8_t right;
    uint8_t top;

    lxw_color_t bottom_color;
    lxw_color_t diag_color;
    lxw_color_t left_color;
    lxw_color_t right_color;
    lxw_color_t top_color;

} lxw_border;




typedef struct lxw_fill {

    lxw_color_t fg_color;
    lxw_color_t bg_color;
    uint8_t pattern;

} lxw_fill;
# 487 "/usr/local/include/xlsxwriter/format.h"
lxw_format *lxw_format_new(void);
void lxw_format_free(lxw_format *format);
int32_t lxw_format_get_xf_index(lxw_format *format);
lxw_font *lxw_format_get_font_key(lxw_format *format);
lxw_border *lxw_format_get_border_key(lxw_format *format);
lxw_fill *lxw_format_get_fill_key(lxw_format *format);
# 514 "/usr/local/include/xlsxwriter/format.h"
void format_set_font_name(lxw_format *format, const char *font_name);
# 534 "/usr/local/include/xlsxwriter/format.h"
void format_set_font_size(lxw_format *format, double size);
# 561 "/usr/local/include/xlsxwriter/format.h"
void format_set_font_color(lxw_format *format, lxw_color_t color);
# 579 "/usr/local/include/xlsxwriter/format.h"
void format_set_bold(lxw_format *format);
# 597 "/usr/local/include/xlsxwriter/format.h"
void format_set_italic(lxw_format *format);
# 621 "/usr/local/include/xlsxwriter/format.h"
void format_set_underline(lxw_format *format, uint8_t style);
# 631 "/usr/local/include/xlsxwriter/format.h"
void format_set_font_strikeout(lxw_format *format);
# 648 "/usr/local/include/xlsxwriter/format.h"
void format_set_font_script(lxw_format *format, uint8_t style);
# 685 "/usr/local/include/xlsxwriter/format.h"
void format_set_num_format(lxw_format *format, const char *num_format);
# 759 "/usr/local/include/xlsxwriter/format.h"
void format_set_num_format_index(lxw_format *format, uint8_t index);
# 785 "/usr/local/include/xlsxwriter/format.h"
void format_set_unlocked(lxw_format *format);
# 809 "/usr/local/include/xlsxwriter/format.h"
void format_set_hidden(lxw_format *format);
# 859 "/usr/local/include/xlsxwriter/format.h"
void format_set_align(lxw_format *format, uint8_t alignment);
# 889 "/usr/local/include/xlsxwriter/format.h"
void format_set_text_wrap(lxw_format *format);
# 912 "/usr/local/include/xlsxwriter/format.h"
void format_set_rotation(lxw_format *format, int16_t angle);
# 941 "/usr/local/include/xlsxwriter/format.h"
void format_set_indent(lxw_format *format, uint8_t level);
# 957 "/usr/local/include/xlsxwriter/format.h"
void format_set_shrink(lxw_format *format);
# 1000 "/usr/local/include/xlsxwriter/format.h"
void format_set_pattern(lxw_format *format, uint8_t index);
# 1029 "/usr/local/include/xlsxwriter/format.h"
void format_set_bg_color(lxw_format *format, lxw_color_t color);
# 1043 "/usr/local/include/xlsxwriter/format.h"
void format_set_fg_color(lxw_format *format, lxw_color_t color);
# 1087 "/usr/local/include/xlsxwriter/format.h"
void format_set_border(lxw_format *format, uint8_t style);
# 1098 "/usr/local/include/xlsxwriter/format.h"
void format_set_bottom(lxw_format *format, uint8_t style);
# 1109 "/usr/local/include/xlsxwriter/format.h"
void format_set_top(lxw_format *format, uint8_t style);
# 1120 "/usr/local/include/xlsxwriter/format.h"
void format_set_left(lxw_format *format, uint8_t style);
# 1131 "/usr/local/include/xlsxwriter/format.h"
void format_set_right(lxw_format *format, uint8_t style);
# 1154 "/usr/local/include/xlsxwriter/format.h"
void format_set_border_color(lxw_format *format, lxw_color_t color);
# 1164 "/usr/local/include/xlsxwriter/format.h"
void format_set_bottom_color(lxw_format *format, lxw_color_t color);
# 1174 "/usr/local/include/xlsxwriter/format.h"
void format_set_top_color(lxw_format *format, lxw_color_t color);
# 1184 "/usr/local/include/xlsxwriter/format.h"
void format_set_left_color(lxw_format *format, lxw_color_t color);
# 1194 "/usr/local/include/xlsxwriter/format.h"
void format_set_right_color(lxw_format *format, lxw_color_t color);

void format_set_diag_type(lxw_format *format, uint8_t value);
void format_set_diag_color(lxw_format *format, lxw_color_t color);
void format_set_diag_border(lxw_format *format, uint8_t value);
void format_set_font_outline(lxw_format *format);
void format_set_font_shadow(lxw_format *format);
void format_set_font_family(lxw_format *format, uint8_t value);
void format_set_font_charset(lxw_format *format, uint8_t value);
void format_set_font_scheme(lxw_format *format, const char *font_scheme);
void format_set_font_condense(lxw_format *format);
void format_set_font_extend(lxw_format *format);
void format_set_reading_order(lxw_format *format, uint8_t value);
void format_set_theme(lxw_format *format, uint8_t value);
void format_set_hyperlink(lxw_format *format);
void format_set_color_indexed(lxw_format *format, uint8_t value);
void format_set_font_only(lxw_format *format);
# 79 "/usr/local/include/xlsxwriter/chart.h" 2

struct lxw_chart_series_list { struct lxw_chart_series *stqh_first; struct lxw_chart_series **stqh_last; };
struct lxw_series_data_points { struct lxw_series_data_point *stqh_first; struct lxw_series_data_point **stqh_last; };







typedef enum lxw_chart_type {


    LXW_CHART_NONE = 0,


    LXW_CHART_AREA,


    LXW_CHART_AREA_STACKED,


    LXW_CHART_AREA_STACKED_PERCENT,


    LXW_CHART_BAR,


    LXW_CHART_BAR_STACKED,


    LXW_CHART_BAR_STACKED_PERCENT,


    LXW_CHART_COLUMN,


    LXW_CHART_COLUMN_STACKED,


    LXW_CHART_COLUMN_STACKED_PERCENT,


    LXW_CHART_DOUGHNUT,


    LXW_CHART_LINE,


    LXW_CHART_PIE,


    LXW_CHART_SCATTER,


    LXW_CHART_SCATTER_STRAIGHT,


    LXW_CHART_SCATTER_STRAIGHT_WITH_MARKERS,


    LXW_CHART_SCATTER_SMOOTH,


    LXW_CHART_SCATTER_SMOOTH_WITH_MARKERS,


    LXW_CHART_RADAR,


    LXW_CHART_RADAR_WITH_MARKERS,


    LXW_CHART_RADAR_FILLED
} lxw_chart_type;




typedef enum lxw_chart_legend_position {


    LXW_CHART_LEGEND_NONE = 0,


    LXW_CHART_LEGEND_RIGHT,


    LXW_CHART_LEGEND_LEFT,


    LXW_CHART_LEGEND_TOP,


    LXW_CHART_LEGEND_BOTTOM,


    LXW_CHART_LEGEND_TOP_RIGHT,


    LXW_CHART_LEGEND_OVERLAY_RIGHT,


    LXW_CHART_LEGEND_OVERLAY_LEFT,


    LXW_CHART_LEGEND_OVERLAY_TOP_RIGHT
} lxw_chart_legend_position;







typedef enum lxw_chart_line_dash_type {


    LXW_CHART_LINE_DASH_SOLID = 0,


    LXW_CHART_LINE_DASH_ROUND_DOT,


    LXW_CHART_LINE_DASH_SQUARE_DOT,


    LXW_CHART_LINE_DASH_DASH,


    LXW_CHART_LINE_DASH_DASH_DOT,


    LXW_CHART_LINE_DASH_LONG_DASH,


    LXW_CHART_LINE_DASH_LONG_DASH_DOT,


    LXW_CHART_LINE_DASH_LONG_DASH_DOT_DOT,


    LXW_CHART_LINE_DASH_DOT,
    LXW_CHART_LINE_DASH_SYSTEM_DASH_DOT,
    LXW_CHART_LINE_DASH_SYSTEM_DASH_DOT_DOT
} lxw_chart_line_dash_type;




typedef enum lxw_chart_marker_type {


    LXW_CHART_MARKER_AUTOMATIC,


    LXW_CHART_MARKER_NONE,


    LXW_CHART_MARKER_SQUARE,


    LXW_CHART_MARKER_DIAMOND,


    LXW_CHART_MARKER_TRIANGLE,


    LXW_CHART_MARKER_X,


    LXW_CHART_MARKER_STAR,


    LXW_CHART_MARKER_SHORT_DASH,


    LXW_CHART_MARKER_LONG_DASH,


    LXW_CHART_MARKER_CIRCLE,


    LXW_CHART_MARKER_PLUS
} lxw_chart_marker_type;




typedef enum lxw_chart_pattern_type {


    LXW_CHART_PATTERN_NONE,


    LXW_CHART_PATTERN_PERCENT_5,


    LXW_CHART_PATTERN_PERCENT_10,


    LXW_CHART_PATTERN_PERCENT_20,


    LXW_CHART_PATTERN_PERCENT_25,


    LXW_CHART_PATTERN_PERCENT_30,


    LXW_CHART_PATTERN_PERCENT_40,


    LXW_CHART_PATTERN_PERCENT_50,


    LXW_CHART_PATTERN_PERCENT_60,


    LXW_CHART_PATTERN_PERCENT_70,


    LXW_CHART_PATTERN_PERCENT_75,


    LXW_CHART_PATTERN_PERCENT_80,


    LXW_CHART_PATTERN_PERCENT_90,


    LXW_CHART_PATTERN_LIGHT_DOWNWARD_DIAGONAL,


    LXW_CHART_PATTERN_LIGHT_UPWARD_DIAGONAL,


    LXW_CHART_PATTERN_DARK_DOWNWARD_DIAGONAL,


    LXW_CHART_PATTERN_DARK_UPWARD_DIAGONAL,


    LXW_CHART_PATTERN_WIDE_DOWNWARD_DIAGONAL,


    LXW_CHART_PATTERN_WIDE_UPWARD_DIAGONAL,


    LXW_CHART_PATTERN_LIGHT_VERTICAL,


    LXW_CHART_PATTERN_LIGHT_HORIZONTAL,


    LXW_CHART_PATTERN_NARROW_VERTICAL,


    LXW_CHART_PATTERN_NARROW_HORIZONTAL,


    LXW_CHART_PATTERN_DARK_VERTICAL,


    LXW_CHART_PATTERN_DARK_HORIZONTAL,


    LXW_CHART_PATTERN_DASHED_DOWNWARD_DIAGONAL,


    LXW_CHART_PATTERN_DASHED_UPWARD_DIAGONAL,


    LXW_CHART_PATTERN_DASHED_HORIZONTAL,


    LXW_CHART_PATTERN_DASHED_VERTICAL,


    LXW_CHART_PATTERN_SMALL_CONFETTI,


    LXW_CHART_PATTERN_LARGE_CONFETTI,


    LXW_CHART_PATTERN_ZIGZAG,


    LXW_CHART_PATTERN_WAVE,


    LXW_CHART_PATTERN_DIAGONAL_BRICK,


    LXW_CHART_PATTERN_HORIZONTAL_BRICK,


    LXW_CHART_PATTERN_WEAVE,


    LXW_CHART_PATTERN_PLAID,


    LXW_CHART_PATTERN_DIVOT,


    LXW_CHART_PATTERN_DOTTED_GRID,


    LXW_CHART_PATTERN_DOTTED_DIAMOND,


    LXW_CHART_PATTERN_SHINGLE,


    LXW_CHART_PATTERN_TRELLIS,


    LXW_CHART_PATTERN_SPHERE,


    LXW_CHART_PATTERN_SMALL_GRID,


    LXW_CHART_PATTERN_LARGE_GRID,


    LXW_CHART_PATTERN_SMALL_CHECK,


    LXW_CHART_PATTERN_LARGE_CHECK,


    LXW_CHART_PATTERN_OUTLINED_DIAMOND,


    LXW_CHART_PATTERN_SOLID_DIAMOND
} lxw_chart_pattern_type;




typedef enum lxw_chart_label_position {

    LXW_CHART_LABEL_POSITION_DEFAULT,


    LXW_CHART_LABEL_POSITION_CENTER,


    LXW_CHART_LABEL_POSITION_RIGHT,


    LXW_CHART_LABEL_POSITION_LEFT,


    LXW_CHART_LABEL_POSITION_ABOVE,


    LXW_CHART_LABEL_POSITION_BELOW,


    LXW_CHART_LABEL_POSITION_INSIDE_BASE,


    LXW_CHART_LABEL_POSITION_INSIDE_END,


    LXW_CHART_LABEL_POSITION_OUTSIDE_END,


    LXW_CHART_LABEL_POSITION_BEST_FIT
} lxw_chart_label_position;




typedef enum lxw_chart_label_separator {

    LXW_CHART_LABEL_SEPARATOR_COMMA,


    LXW_CHART_LABEL_SEPARATOR_SEMICOLON,


    LXW_CHART_LABEL_SEPARATOR_PERIOD,


    LXW_CHART_LABEL_SEPARATOR_NEWLINE,


    LXW_CHART_LABEL_SEPARATOR_SPACE
} lxw_chart_label_separator;




typedef enum lxw_chart_axis_type {

    LXW_CHART_AXIS_TYPE_X,


    LXW_CHART_AXIS_TYPE_Y
} lxw_chart_axis_type;

enum lxw_chart_subtype {

    LXW_CHART_SUBTYPE_NONE = 0,
    LXW_CHART_SUBTYPE_STACKED,
    LXW_CHART_SUBTYPE_STACKED_PERCENT
};

enum lxw_chart_grouping {
    LXW_GROUPING_CLUSTERED,
    LXW_GROUPING_STANDARD,
    LXW_GROUPING_PERCENTSTACKED,
    LXW_GROUPING_STACKED
};




typedef enum lxw_chart_axis_tick_position {

    LXW_CHART_AXIS_POSITION_DEFAULT,


    LXW_CHART_AXIS_POSITION_ON_TICK,


    LXW_CHART_AXIS_POSITION_BETWEEN
} lxw_chart_axis_tick_position;




typedef enum lxw_chart_axis_label_position {


    LXW_CHART_AXIS_LABEL_POSITION_NEXT_TO,



    LXW_CHART_AXIS_LABEL_POSITION_HIGH,



    LXW_CHART_AXIS_LABEL_POSITION_LOW,


    LXW_CHART_AXIS_LABEL_POSITION_NONE
} lxw_chart_axis_label_position;




typedef enum lxw_chart_axis_label_alignment {

    LXW_CHART_AXIS_LABEL_ALIGN_CENTER,


    LXW_CHART_AXIS_LABEL_ALIGN_LEFT,


    LXW_CHART_AXIS_LABEL_ALIGN_RIGHT
} lxw_chart_axis_label_alignment;




typedef enum lxw_chart_axis_display_unit {


    LXW_CHART_AXIS_UNITS_NONE,


    LXW_CHART_AXIS_UNITS_HUNDREDS,


    LXW_CHART_AXIS_UNITS_THOUSANDS,


    LXW_CHART_AXIS_UNITS_TEN_THOUSANDS,


    LXW_CHART_AXIS_UNITS_HUNDRED_THOUSANDS,


    LXW_CHART_AXIS_UNITS_MILLIONS,


    LXW_CHART_AXIS_UNITS_TEN_MILLIONS,


    LXW_CHART_AXIS_UNITS_HUNDRED_MILLIONS,


    LXW_CHART_AXIS_UNITS_BILLIONS,


    LXW_CHART_AXIS_UNITS_TRILLIONS
} lxw_chart_axis_display_unit;




typedef enum lxw_chart_axis_tick_mark {


    LXW_CHART_AXIS_TICK_MARK_DEFAULT,


    LXW_CHART_AXIS_TICK_MARK_NONE,


    LXW_CHART_AXIS_TICK_MARK_INSIDE,


    LXW_CHART_AXIS_TICK_MARK_OUTSIDE,


    LXW_CHART_AXIS_TICK_MARK_CROSSING
} lxw_chart_tick_mark;

typedef struct lxw_series_range {
    char *formula;
    char *sheetname;
    lxw_row_t first_row;
    lxw_row_t last_row;
    lxw_col_t first_col;
    lxw_col_t last_col;
    uint8_t ignore_cache;

    uint8_t has_string_cache;
    uint16_t num_data_points;
    struct lxw_series_data_points *data_cache;

} lxw_series_range;

typedef struct lxw_series_data_point {
    uint8_t is_string;
    double number;
    char *string;
    uint8_t no_data;

    struct { struct lxw_series_data_point *stqe_next; } list_pointers;

} lxw_series_data_point;






typedef struct lxw_chart_line {


    lxw_color_t color;


    uint8_t none;


    float width;


    uint8_t dash_type;


    uint8_t transparency;

} lxw_chart_line;






typedef struct lxw_chart_fill {


    lxw_color_t color;


    uint8_t none;


    uint8_t transparency;

} lxw_chart_fill;






typedef struct lxw_chart_pattern {


    lxw_color_t fg_color;


    lxw_color_t bg_color;


    uint8_t type;

} lxw_chart_pattern;






typedef struct lxw_chart_font {


    char *name;


    double size;


    uint8_t bold;


    uint8_t italic;


    uint8_t underline;





    int32_t rotation;


    lxw_color_t color;


    uint8_t pitch_family;


    uint8_t charset;


    int8_t baseline;

} lxw_chart_font;

typedef struct lxw_chart_marker {

    uint8_t type;
    uint8_t size;
    lxw_chart_line *line;
    lxw_chart_fill *fill;
    lxw_chart_pattern *pattern;

} lxw_chart_marker;

typedef struct lxw_chart_legend {

    lxw_chart_font *font;
    uint8_t position;

} lxw_chart_legend;

typedef struct lxw_chart_title {

    char *name;
    lxw_row_t row;
    lxw_col_t col;
    lxw_chart_font *font;
    uint8_t off;
    uint8_t is_horizontal;
    uint8_t ignore_cache;



    lxw_series_range *range;

    struct lxw_series_data_point data_point;

} lxw_chart_title;







typedef struct lxw_chart_point {


    lxw_chart_line *line;


    lxw_chart_fill *fill;


    lxw_chart_pattern *pattern;

} lxw_chart_point;




typedef enum lxw_chart_blank {


    LXW_CHART_BLANKS_AS_GAP,


    LXW_CHART_BLANKS_AS_ZERO,


    LXW_CHART_BLANKS_AS_CONNECTED
} lxw_chart_blank;

enum lxw_chart_position {
    LXW_CHART_AXIS_RIGHT,
    LXW_CHART_AXIS_LEFT,
    LXW_CHART_AXIS_TOP,
    LXW_CHART_AXIS_BOTTOM
};




typedef enum lxw_chart_error_bar_type {

    LXW_CHART_ERROR_BAR_TYPE_STD_ERROR,


    LXW_CHART_ERROR_BAR_TYPE_FIXED,


    LXW_CHART_ERROR_BAR_TYPE_PERCENTAGE,


    LXW_CHART_ERROR_BAR_TYPE_STD_DEV
} lxw_chart_error_bar_type;




typedef enum lxw_chart_error_bar_direction {


    LXW_CHART_ERROR_BAR_DIR_BOTH,


    LXW_CHART_ERROR_BAR_DIR_PLUS,


    LXW_CHART_ERROR_BAR_DIR_MINUS
} lxw_chart_error_bar_direction;




typedef enum lxw_chart_error_bar_axis {

    LXW_CHART_ERROR_BAR_AXIS_X,


    LXW_CHART_ERROR_BAR_AXIS_Y
} lxw_chart_error_bar_axis;




typedef enum lxw_chart_error_bar_cap {

    LXW_CHART_ERROR_BAR_END_CAP,


    LXW_CHART_ERROR_BAR_NO_CAP
} lxw_chart_error_bar_cap;

typedef struct lxw_series_error_bars {
    uint8_t type;
    uint8_t direction;
    uint8_t endcap;
    uint8_t has_value;
    uint8_t is_set;
    uint8_t is_x;
    uint8_t chart_group;
    double value;
    lxw_chart_line *line;

} lxw_series_error_bars;




typedef enum lxw_chart_trendline_type {

    LXW_CHART_TRENDLINE_TYPE_LINEAR,


    LXW_CHART_TRENDLINE_TYPE_LOG,


    LXW_CHART_TRENDLINE_TYPE_POLY,


    LXW_CHART_TRENDLINE_TYPE_POWER,


    LXW_CHART_TRENDLINE_TYPE_EXP,


    LXW_CHART_TRENDLINE_TYPE_AVERAGE
} lxw_chart_trendline_type;
# 903 "/usr/local/include/xlsxwriter/chart.h"
typedef struct lxw_chart_series {

    lxw_series_range *categories;
    lxw_series_range *values;
    lxw_chart_title title;
    lxw_chart_line *line;
    lxw_chart_fill *fill;
    lxw_chart_pattern *pattern;
    lxw_chart_marker *marker;
    lxw_chart_point *points;
    uint16_t point_count;

    uint8_t smooth;
    uint8_t invert_if_negative;


    uint8_t has_labels;
    uint8_t show_labels_value;
    uint8_t show_labels_category;
    uint8_t show_labels_name;
    uint8_t show_labels_leader;
    uint8_t show_labels_legend;
    uint8_t show_labels_percent;
    uint8_t label_position;
    uint8_t label_separator;
    uint8_t default_label_position;
    char *label_num_format;
    lxw_chart_font *label_font;

    lxw_series_error_bars *x_error_bars;
    lxw_series_error_bars *y_error_bars;

    uint8_t has_trendline;
    uint8_t has_trendline_forecast;
    uint8_t has_trendline_equation;
    uint8_t has_trendline_r_squared;
    uint8_t has_trendline_intercept;
    uint8_t trendline_type;
    uint8_t trendline_value;
    double trendline_forward;
    double trendline_backward;
    uint8_t trendline_value_type;
    char *trendline_name;
    lxw_chart_line *trendline_line;
    double trendline_intercept;

    struct { struct lxw_chart_series *stqe_next; } list_pointers;

} lxw_chart_series;


typedef struct lxw_chart_gridline {

    uint8_t visible;
    lxw_chart_line *line;

} lxw_chart_gridline;







typedef struct lxw_chart_axis {

    lxw_chart_title title;

    char *num_format;
    char *default_num_format;
    uint8_t source_linked;

    uint8_t major_tick_mark;
    uint8_t minor_tick_mark;
    uint8_t is_horizontal;

    lxw_chart_gridline major_gridlines;
    lxw_chart_gridline minor_gridlines;

    lxw_chart_font *num_font;
    lxw_chart_line *line;
    lxw_chart_fill *fill;
    lxw_chart_pattern *pattern;

    uint8_t is_category;
    uint8_t is_date;
    uint8_t is_value;
    uint8_t axis_position;
    uint8_t position_axis;
    uint8_t label_position;
    uint8_t label_align;
    uint8_t hidden;
    uint8_t reverse;

    uint8_t has_min;
    double min;
    uint8_t has_max;
    double max;

    uint8_t has_major_unit;
    double major_unit;
    uint8_t has_minor_unit;
    double minor_unit;

    uint16_t interval_unit;
    uint16_t interval_tick;

    uint16_t log_base;

    uint8_t display_units;
    uint8_t display_units_visible;

    uint8_t has_crossing;
    uint8_t crossing_max;
    double crossing;

} lxw_chart_axis;







typedef struct lxw_chart {

    FILE *file;

    uint8_t type;
    uint8_t subtype;
    uint16_t series_index;

    void (*write_chart_type) (struct lxw_chart *);
    void (*write_plot_area) (struct lxw_chart *);





    lxw_chart_axis *x_axis;





    lxw_chart_axis *y_axis;

    lxw_chart_title title;

    uint32_t id;
    uint32_t axis_id_1;
    uint32_t axis_id_2;
    uint32_t axis_id_3;
    uint32_t axis_id_4;

    uint8_t in_use;
    uint8_t chart_group;
    uint8_t cat_has_num_fmt;
    uint8_t is_chartsheet;

    uint8_t has_horiz_cat_axis;
    uint8_t has_horiz_val_axis;

    uint8_t style_id;
    uint16_t rotation;
    uint16_t hole_size;

    uint8_t no_title;
    uint8_t has_overlap;
    int8_t overlap_y1;
    int8_t overlap_y2;
    uint16_t gap_y1;
    uint16_t gap_y2;

    uint8_t grouping;
    uint8_t default_cross_between;

    lxw_chart_legend legend;
    int16_t *delete_series;
    uint16_t delete_series_count;
    lxw_chart_marker *default_marker;

    lxw_chart_line *chartarea_line;
    lxw_chart_fill *chartarea_fill;
    lxw_chart_pattern *chartarea_pattern;
    lxw_chart_line *plotarea_line;
    lxw_chart_fill *plotarea_fill;
    lxw_chart_pattern *plotarea_pattern;

    uint8_t has_drop_lines;
    lxw_chart_line *drop_lines_line;

    uint8_t has_high_low_lines;
    lxw_chart_line *high_low_lines_line;

    struct lxw_chart_series_list *series_list;

    uint8_t has_table;
    uint8_t has_table_vertical;
    uint8_t has_table_horizontal;
    uint8_t has_table_outline;
    uint8_t has_table_legend_keys;
    lxw_chart_font *table_font;

    uint8_t show_blanks_as;
    uint8_t show_hidden_data;

    uint8_t has_up_down_bars;
    lxw_chart_line *up_bar_line;
    lxw_chart_line *down_bar_line;
    lxw_chart_fill *up_bar_fill;
    lxw_chart_fill *down_bar_fill;

    uint8_t default_label_position;
    uint8_t is_protected;

    struct { struct lxw_chart *stqe_next; } ordered_list_pointers;
    struct { struct lxw_chart *stqe_next; } list_pointers;

} lxw_chart;
# 1131 "/usr/local/include/xlsxwriter/chart.h"
lxw_chart *lxw_chart_new(uint8_t type);
void lxw_chart_free(lxw_chart *chart);
void lxw_chart_assemble_xml_file(lxw_chart *chart);
# 1213 "/usr/local/include/xlsxwriter/chart.h"
lxw_chart_series *chart_add_series(lxw_chart *chart,
                                   const char *categories,
                                   const char *values);
# 1246 "/usr/local/include/xlsxwriter/chart.h"
void chart_series_set_categories(lxw_chart_series *series,
                                 const char *sheetname, lxw_row_t first_row,
                                 lxw_col_t first_col, lxw_row_t last_row,
                                 lxw_col_t last_col);
# 1269 "/usr/local/include/xlsxwriter/chart.h"
void chart_series_set_values(lxw_chart_series *series, const char *sheetname,
                             lxw_row_t first_row, lxw_col_t first_col,
                             lxw_row_t last_row, lxw_col_t last_col);
# 1305 "/usr/local/include/xlsxwriter/chart.h"
void chart_series_set_name(lxw_chart_series *series, const char *name);
# 1325 "/usr/local/include/xlsxwriter/chart.h"
void chart_series_set_name_range(lxw_chart_series *series,
                                 const char *sheetname, lxw_row_t row,
                                 lxw_col_t col);
# 1348 "/usr/local/include/xlsxwriter/chart.h"
void chart_series_set_line(lxw_chart_series *series, lxw_chart_line *line);
# 1372 "/usr/local/include/xlsxwriter/chart.h"
void chart_series_set_fill(lxw_chart_series *series, lxw_chart_fill *fill);
# 1387 "/usr/local/include/xlsxwriter/chart.h"
void chart_series_set_invert_if_negative(lxw_chart_series *series);
# 1415 "/usr/local/include/xlsxwriter/chart.h"
void chart_series_set_pattern(lxw_chart_series *series,
                              lxw_chart_pattern *pattern);
# 1478 "/usr/local/include/xlsxwriter/chart.h"
void chart_series_set_marker_type(lxw_chart_series *series, uint8_t type);
# 1497 "/usr/local/include/xlsxwriter/chart.h"
void chart_series_set_marker_size(lxw_chart_series *series, uint8_t size);
# 1522 "/usr/local/include/xlsxwriter/chart.h"
void chart_series_set_marker_line(lxw_chart_series *series,
                                  lxw_chart_line *line);
# 1539 "/usr/local/include/xlsxwriter/chart.h"
void chart_series_set_marker_fill(lxw_chart_series *series,
                                  lxw_chart_fill *fill);
# 1556 "/usr/local/include/xlsxwriter/chart.h"
void chart_series_set_marker_pattern(lxw_chart_series *series,
                                     lxw_chart_pattern *pattern);
# 1583 "/usr/local/include/xlsxwriter/chart.h"
lxw_error chart_series_set_points(lxw_chart_series *series,
                                  lxw_chart_point *points[]);
# 1604 "/usr/local/include/xlsxwriter/chart.h"
void chart_series_set_smooth(lxw_chart_series *series, uint8_t smooth);
# 1630 "/usr/local/include/xlsxwriter/chart.h"
void chart_series_set_labels(lxw_chart_series *series);
# 1652 "/usr/local/include/xlsxwriter/chart.h"
void chart_series_set_labels_options(lxw_chart_series *series,
                                     uint8_t show_name, uint8_t show_category,
                                     uint8_t show_value);
# 1686 "/usr/local/include/xlsxwriter/chart.h"
void chart_series_set_labels_separator(lxw_chart_series *series,
                                       uint8_t separator);
# 1723 "/usr/local/include/xlsxwriter/chart.h"
void chart_series_set_labels_position(lxw_chart_series *series,
                                      uint8_t position);
# 1748 "/usr/local/include/xlsxwriter/chart.h"
void chart_series_set_labels_leader_line(lxw_chart_series *series);
# 1767 "/usr/local/include/xlsxwriter/chart.h"
void chart_series_set_labels_legend(lxw_chart_series *series);
# 1788 "/usr/local/include/xlsxwriter/chart.h"
void chart_series_set_labels_percentage(lxw_chart_series *series);
# 1811 "/usr/local/include/xlsxwriter/chart.h"
void chart_series_set_labels_num_format(lxw_chart_series *series,
                                        const char *num_format);
# 1836 "/usr/local/include/xlsxwriter/chart.h"
void chart_series_set_labels_font(lxw_chart_series *series,
                                  lxw_chart_font *font);
# 1893 "/usr/local/include/xlsxwriter/chart.h"
void chart_series_set_trendline(lxw_chart_series *series, uint8_t type,
                                uint8_t value);
# 1917 "/usr/local/include/xlsxwriter/chart.h"
void chart_series_set_trendline_forecast(lxw_chart_series *series,
                                         double forward, double backward);
# 1939 "/usr/local/include/xlsxwriter/chart.h"
void chart_series_set_trendline_equation(lxw_chart_series *series);
# 1960 "/usr/local/include/xlsxwriter/chart.h"
void chart_series_set_trendline_r_squared(lxw_chart_series *series);
# 1987 "/usr/local/include/xlsxwriter/chart.h"
void chart_series_set_trendline_intercept(lxw_chart_series *series,
                                          double intercept);
# 2027 "/usr/local/include/xlsxwriter/chart.h"
void chart_series_set_trendline_name(lxw_chart_series *series,
                                     const char *name);
# 2051 "/usr/local/include/xlsxwriter/chart.h"
void chart_series_set_trendline_line(lxw_chart_series *series,
                                     lxw_chart_line *line);
# 2091 "/usr/local/include/xlsxwriter/chart.h"
lxw_series_error_bars *chart_series_get_error_bars(lxw_chart_series *series, lxw_chart_error_bar_axis
                                                   axis_type);
# 2148 "/usr/local/include/xlsxwriter/chart.h"
void chart_series_set_error_bars(lxw_series_error_bars *error_bars,
                                 uint8_t type, double value);
# 2180 "/usr/local/include/xlsxwriter/chart.h"
void chart_series_set_error_bars_direction(lxw_series_error_bars *error_bars,
                                           uint8_t direction);
# 2209 "/usr/local/include/xlsxwriter/chart.h"
void chart_series_set_error_bars_endcap(lxw_series_error_bars *error_bars,
                                        uint8_t endcap);
# 2235 "/usr/local/include/xlsxwriter/chart.h"
void chart_series_set_error_bars_line(lxw_series_error_bars *error_bars,
                                      lxw_chart_line *line);
# 2266 "/usr/local/include/xlsxwriter/chart.h"
lxw_chart_axis *chart_axis_get(lxw_chart *chart,
                               lxw_chart_axis_type axis_type);
# 2299 "/usr/local/include/xlsxwriter/chart.h"
void chart_axis_set_name(lxw_chart_axis *axis, const char *name);
# 2321 "/usr/local/include/xlsxwriter/chart.h"
void chart_axis_set_name_range(lxw_chart_axis *axis, const char *sheetname,
                               lxw_row_t row, lxw_col_t col);
# 2347 "/usr/local/include/xlsxwriter/chart.h"
void chart_axis_set_name_font(lxw_chart_axis *axis, lxw_chart_font *font);
# 2371 "/usr/local/include/xlsxwriter/chart.h"
void chart_axis_set_num_font(lxw_chart_axis *axis, lxw_chart_font *font);
# 2395 "/usr/local/include/xlsxwriter/chart.h"
void chart_axis_set_num_format(lxw_chart_axis *axis, const char *num_format);
# 2419 "/usr/local/include/xlsxwriter/chart.h"
void chart_axis_set_line(lxw_chart_axis *axis, lxw_chart_line *line);
# 2442 "/usr/local/include/xlsxwriter/chart.h"
void chart_axis_set_fill(lxw_chart_axis *axis, lxw_chart_fill *fill);
# 2461 "/usr/local/include/xlsxwriter/chart.h"
void chart_axis_set_pattern(lxw_chart_axis *axis, lxw_chart_pattern *pattern);
# 2479 "/usr/local/include/xlsxwriter/chart.h"
void chart_axis_set_reverse(lxw_chart_axis *axis);
# 2502 "/usr/local/include/xlsxwriter/chart.h"
void chart_axis_set_crossing(lxw_chart_axis *axis, double value);
# 2524 "/usr/local/include/xlsxwriter/chart.h"
void chart_axis_set_crossing_max(lxw_chart_axis *axis);
# 2542 "/usr/local/include/xlsxwriter/chart.h"
void chart_axis_off(lxw_chart_axis *axis);
# 2566 "/usr/local/include/xlsxwriter/chart.h"
void chart_axis_set_position(lxw_chart_axis *axis, uint8_t position);
# 2603 "/usr/local/include/xlsxwriter/chart.h"
void chart_axis_set_label_position(lxw_chart_axis *axis, uint8_t position);
# 2628 "/usr/local/include/xlsxwriter/chart.h"
void chart_axis_set_label_align(lxw_chart_axis *axis, uint8_t align);
# 2648 "/usr/local/include/xlsxwriter/chart.h"
void chart_axis_set_min(lxw_chart_axis *axis, double min);
# 2668 "/usr/local/include/xlsxwriter/chart.h"
void chart_axis_set_max(lxw_chart_axis *axis, double max);
# 2690 "/usr/local/include/xlsxwriter/chart.h"
void chart_axis_set_log_base(lxw_chart_axis *axis, uint16_t log_base);
# 2723 "/usr/local/include/xlsxwriter/chart.h"
void chart_axis_set_major_tick_mark(lxw_chart_axis *axis, uint8_t type);
# 2742 "/usr/local/include/xlsxwriter/chart.h"
void chart_axis_set_minor_tick_mark(lxw_chart_axis *axis, uint8_t type);
# 2770 "/usr/local/include/xlsxwriter/chart.h"
void chart_axis_set_interval_unit(lxw_chart_axis *axis, uint16_t unit);
# 2790 "/usr/local/include/xlsxwriter/chart.h"
void chart_axis_set_interval_tick(lxw_chart_axis *axis, uint16_t unit);
# 2813 "/usr/local/include/xlsxwriter/chart.h"
void chart_axis_set_major_unit(lxw_chart_axis *axis, double unit);
# 2832 "/usr/local/include/xlsxwriter/chart.h"
void chart_axis_set_minor_unit(lxw_chart_axis *axis, double unit);
# 2853 "/usr/local/include/xlsxwriter/chart.h"
void chart_axis_set_display_units(lxw_chart_axis *axis, uint8_t units);
# 2871 "/usr/local/include/xlsxwriter/chart.h"
void chart_axis_set_display_units_visible(lxw_chart_axis *axis,
                                          uint8_t visible);
# 2897 "/usr/local/include/xlsxwriter/chart.h"
void chart_axis_major_gridlines_set_visible(lxw_chart_axis *axis,
                                            uint8_t visible);
# 2923 "/usr/local/include/xlsxwriter/chart.h"
void chart_axis_minor_gridlines_set_visible(lxw_chart_axis *axis,
                                            uint8_t visible);
# 2959 "/usr/local/include/xlsxwriter/chart.h"
void chart_axis_major_gridlines_set_line(lxw_chart_axis *axis,
                                         lxw_chart_line *line);
# 2976 "/usr/local/include/xlsxwriter/chart.h"
void chart_axis_minor_gridlines_set_line(lxw_chart_axis *axis,
                                         lxw_chart_line *line);
# 3006 "/usr/local/include/xlsxwriter/chart.h"
void chart_title_set_name(lxw_chart *chart, const char *name);
# 3024 "/usr/local/include/xlsxwriter/chart.h"
void chart_title_set_name_range(lxw_chart *chart, const char *sheetname,
                                lxw_row_t row, lxw_col_t col);
# 3047 "/usr/local/include/xlsxwriter/chart.h"
void chart_title_set_name_font(lxw_chart *chart, lxw_chart_font *font);
# 3064 "/usr/local/include/xlsxwriter/chart.h"
void chart_title_off(lxw_chart *chart);
# 3102 "/usr/local/include/xlsxwriter/chart.h"
void chart_legend_set_position(lxw_chart *chart, uint8_t position);
# 3123 "/usr/local/include/xlsxwriter/chart.h"
void chart_legend_set_font(lxw_chart *chart, lxw_chart_font *font);
# 3150 "/usr/local/include/xlsxwriter/chart.h"
lxw_error chart_legend_delete_series(lxw_chart *chart,
                                     int16_t delete_series[]);
# 3174 "/usr/local/include/xlsxwriter/chart.h"
void chart_chartarea_set_line(lxw_chart *chart, lxw_chart_line *line);
# 3192 "/usr/local/include/xlsxwriter/chart.h"
void chart_chartarea_set_fill(lxw_chart *chart, lxw_chart_fill *fill);
# 3208 "/usr/local/include/xlsxwriter/chart.h"
void chart_chartarea_set_pattern(lxw_chart *chart,
                                 lxw_chart_pattern *pattern);
# 3235 "/usr/local/include/xlsxwriter/chart.h"
void chart_plotarea_set_line(lxw_chart *chart, lxw_chart_line *line);
# 3253 "/usr/local/include/xlsxwriter/chart.h"
void chart_plotarea_set_fill(lxw_chart *chart, lxw_chart_fill *fill);
# 3269 "/usr/local/include/xlsxwriter/chart.h"
void chart_plotarea_set_pattern(lxw_chart *chart, lxw_chart_pattern *pattern);
# 3298 "/usr/local/include/xlsxwriter/chart.h"
void chart_set_style(lxw_chart *chart, uint8_t style_id);
# 3318 "/usr/local/include/xlsxwriter/chart.h"
void chart_set_table(lxw_chart *chart);
# 3362 "/usr/local/include/xlsxwriter/chart.h"
void chart_set_table_grid(lxw_chart *chart, uint8_t horizontal,
                          uint8_t vertical, uint8_t outline,
                          uint8_t legend_keys);

void chart_set_table_font(lxw_chart *chart, lxw_chart_font *font);
# 3386 "/usr/local/include/xlsxwriter/chart.h"
void chart_set_up_down_bars(lxw_chart *chart);
# 3414 "/usr/local/include/xlsxwriter/chart.h"
void chart_set_up_down_bars_format(lxw_chart *chart,
                                   lxw_chart_line *up_bar_line,
                                   lxw_chart_fill *up_bar_fill,
                                   lxw_chart_line *down_bar_line,
                                   lxw_chart_fill *down_bar_fill);
# 3447 "/usr/local/include/xlsxwriter/chart.h"
void chart_set_drop_lines(lxw_chart *chart, lxw_chart_line *line);
# 3476 "/usr/local/include/xlsxwriter/chart.h"
void chart_set_high_low_lines(lxw_chart *chart, lxw_chart_line *line);
# 3498 "/usr/local/include/xlsxwriter/chart.h"
void chart_set_series_overlap(lxw_chart *chart, int8_t overlap);
# 3520 "/usr/local/include/xlsxwriter/chart.h"
void chart_set_series_gap(lxw_chart *chart, uint16_t gap);
# 3543 "/usr/local/include/xlsxwriter/chart.h"
void chart_show_blanks_as(lxw_chart *chart, uint8_t option);
# 3556 "/usr/local/include/xlsxwriter/chart.h"
void chart_show_hidden_data(lxw_chart *chart);
# 3577 "/usr/local/include/xlsxwriter/chart.h"
void chart_set_rotation(lxw_chart *chart, uint16_t rotation);
# 3597 "/usr/local/include/xlsxwriter/chart.h"
void chart_set_hole_size(lxw_chart *chart, uint8_t size);

lxw_error lxw_chart_add_data_cache(lxw_series_range *range, uint8_t *data,
                                   uint16_t rows, uint8_t cols, uint8_t col);
# 53 "/usr/local/include/xlsxwriter/worksheet.h" 2
# 1 "/usr/local/include/xlsxwriter/drawing.h" 1
# 17 "/usr/local/include/xlsxwriter/drawing.h"
struct lxw_drawing_objects { struct lxw_drawing_object *stqh_first; struct lxw_drawing_object **stqh_last; };

enum lxw_drawing_types {
    LXW_DRAWING_NONE = 0,
    LXW_DRAWING_IMAGE,
    LXW_DRAWING_CHART,
    LXW_DRAWING_SHAPE
};

enum image_types {
    LXW_IMAGE_UNKNOWN = 0,
    LXW_IMAGE_PNG,
    LXW_IMAGE_JPEG,
    LXW_IMAGE_BMP
};


typedef struct lxw_drawing_coords {
    uint32_t col;
    uint32_t row;
    double col_offset;
    double row_offset;
} lxw_drawing_coords;


typedef struct lxw_drawing_object {
    uint8_t type;
    uint8_t anchor;
    struct lxw_drawing_coords from;
    struct lxw_drawing_coords to;
    uint32_t col_absolute;
    uint32_t row_absolute;
    uint32_t width;
    uint32_t height;
    uint8_t shape;
    uint32_t rel_index;
    uint32_t url_rel_index;
    char *description;
    char *tip;

    struct { struct lxw_drawing_object *stqe_next; } list_pointers;

} lxw_drawing_object;




typedef struct lxw_drawing {

    FILE *file;

    uint8_t embedded;
    uint8_t orientation;

    struct lxw_drawing_objects *drawing_objects;

} lxw_drawing;
# 82 "/usr/local/include/xlsxwriter/drawing.h"
lxw_drawing *lxw_drawing_new(void);
void lxw_drawing_free(lxw_drawing *drawing);
void lxw_drawing_assemble_xml_file(lxw_drawing *self);
void lxw_free_drawing_object(struct lxw_drawing_object *drawing_object);
void lxw_add_drawing_object(lxw_drawing *drawing,
                            lxw_drawing_object *drawing_object);
# 54 "/usr/local/include/xlsxwriter/worksheet.h" 2


# 1 "/usr/local/include/xlsxwriter/styles.h" 1
# 13 "/usr/local/include/xlsxwriter/styles.h"
# 1 "/app/src/headers/override/ctype.h" 1
# 14 "/usr/local/include/xlsxwriter/styles.h" 2






typedef struct lxw_styles {

    FILE *file;
    uint32_t font_count;
    uint32_t xf_count;
    uint32_t dxf_count;
    uint32_t num_format_count;
    uint32_t border_count;
    uint32_t fill_count;
    struct lxw_formats *xf_formats;
    struct lxw_formats *dxf_formats;
    uint8_t has_hyperlink;
    uint16_t hyperlink_font_id;
    uint8_t has_comments;

} lxw_styles;
# 44 "/usr/local/include/xlsxwriter/styles.h"
lxw_styles *lxw_styles_new(void);
void lxw_styles_free(lxw_styles *styles);
void lxw_styles_assemble_xml_file(lxw_styles *self);
void lxw_styles_write_string_fragment(lxw_styles *self, char *string);
void lxw_styles_write_rich_font(lxw_styles *lxw_styles, lxw_format *format);
# 57 "/usr/local/include/xlsxwriter/worksheet.h" 2
# 1 "/usr/local/include/xlsxwriter/utility.h" 1
# 24 "/usr/local/include/xlsxwriter/utility.h"
# 1 "/usr/local/include/xlsxwriter/xmlwriter.h" 1
# 25 "/usr/local/include/xlsxwriter/xmlwriter.h"
# 1 "/usr/local/include/xlsxwriter/utility.h" 1
# 26 "/usr/local/include/xlsxwriter/xmlwriter.h" 2
# 44 "/usr/local/include/xlsxwriter/xmlwriter.h"
struct xml_attribute {
    char key[2080];
    char value[2080];


    struct { struct xml_attribute *stqe_next; } list_entries;
};


struct xml_attribute_list { struct xml_attribute *stqh_first; struct xml_attribute **stqh_last; };


struct xml_attribute *lxw_new_attribute_str(const char *key,
                                            const char *value);
struct xml_attribute *lxw_new_attribute_int(const char *key, uint32_t value);
struct xml_attribute *lxw_new_attribute_dbl(const char *key, double value);
# 99 "/usr/local/include/xlsxwriter/xmlwriter.h"
void lxw_xml_declaration(FILE * xmlfile);
# 108 "/usr/local/include/xlsxwriter/xmlwriter.h"
void lxw_xml_start_tag(FILE * xmlfile,
                       const char *tag,
                       struct xml_attribute_list *attributes);
# 120 "/usr/local/include/xlsxwriter/xmlwriter.h"
void lxw_xml_start_tag_unencoded(FILE * xmlfile,
                                 const char *tag,
                                 struct xml_attribute_list *attributes);







void lxw_xml_end_tag(FILE * xmlfile, const char *tag);
# 139 "/usr/local/include/xlsxwriter/xmlwriter.h"
void lxw_xml_empty_tag(FILE * xmlfile,
                       const char *tag,
                       struct xml_attribute_list *attributes);
# 151 "/usr/local/include/xlsxwriter/xmlwriter.h"
void lxw_xml_empty_tag_unencoded(FILE * xmlfile,
                                 const char *tag,
                                 struct xml_attribute_list *attributes);
# 163 "/usr/local/include/xlsxwriter/xmlwriter.h"
void lxw_xml_data_element(FILE * xmlfile,
                          const char *tag,
                          const char *data,
                          struct xml_attribute_list *attributes);

void lxw_xml_rich_si_element(FILE * xmlfile, const char *string);

char *lxw_escape_control_characters(const char *string);
char *lxw_escape_url_characters(const char *string, uint8_t escape_hash);

char *lxw_escape_data(const char *data);
# 25 "/usr/local/include/xlsxwriter/utility.h" 2
# 104 "/usr/local/include/xlsxwriter/utility.h"
const char *lxw_version(void);
# 118 "/usr/local/include/xlsxwriter/utility.h"
uint16_t lxw_version_id(void);
# 147 "/usr/local/include/xlsxwriter/utility.h"
char *lxw_strerror(lxw_error error_num);


char *lxw_quote_sheetname(const char *str);

void lxw_col_to_name(char *col_name, lxw_col_t col_num, uint8_t absolute);

void lxw_rowcol_to_cell(char *cell_name, lxw_row_t row, lxw_col_t col);

void lxw_rowcol_to_cell_abs(char *cell_name,
                            lxw_row_t row,
                            lxw_col_t col, uint8_t abs_row, uint8_t abs_col);

void lxw_rowcol_to_range(char *range,
                         lxw_row_t first_row, lxw_col_t first_col,
                         lxw_row_t last_row, lxw_col_t last_col);

void lxw_rowcol_to_range_abs(char *range,
                             lxw_row_t first_row, lxw_col_t first_col,
                             lxw_row_t last_row, lxw_col_t last_col);

void lxw_rowcol_to_formula_abs(char *formula, const char *sheetname,
                               lxw_row_t first_row, lxw_col_t first_col,
                               lxw_row_t last_row, lxw_col_t last_col);

uint32_t lxw_name_to_row(const char *row_str);
uint16_t lxw_name_to_col(const char *col_str);
uint32_t lxw_name_to_row_2(const char *row_str);
uint16_t lxw_name_to_col_2(const char *col_str);

double lxw_datetime_to_excel_date(lxw_datetime *datetime, uint8_t date_1904);

char *lxw_strdup(const char *str);
char *lxw_strdup_formula(const char *formula);

size_t lxw_utf8_strlen(const char *str);

void lxw_str_tolower(char *str);
# 193 "/usr/local/include/xlsxwriter/utility.h"
FILE *lxw_tmpfile(char *tmpdir);
FILE *lxw_fopen(const char *filename, const char *mode);
# 205 "/usr/local/include/xlsxwriter/utility.h"
uint16_t lxw_hash_password(const char *password);
# 58 "/usr/local/include/xlsxwriter/worksheet.h" 2
# 1 "/usr/local/include/xlsxwriter/relationships.h" 1
# 18 "/usr/local/include/xlsxwriter/relationships.h"
struct lxw_rel_tuples { struct lxw_rel_tuple *stqh_first; struct lxw_rel_tuple **stqh_last; };

typedef struct lxw_rel_tuple {

    char *type;
    char *target;
    char *target_mode;

    struct { struct lxw_rel_tuple *stqe_next; } list_pointers;

} lxw_rel_tuple;




typedef struct lxw_relationships {

    FILE *file;

    uint32_t rel_id;
    struct lxw_rel_tuples *relationships;

} lxw_relationships;
# 50 "/usr/local/include/xlsxwriter/relationships.h"
lxw_relationships *lxw_relationships_new(void);
void lxw_free_relationships(lxw_relationships *relationships);
void lxw_relationships_assemble_xml_file(lxw_relationships *self);

void lxw_add_document_relationship(lxw_relationships *self, const char *type,
                                   const char *target);
void lxw_add_package_relationship(lxw_relationships *self, const char *type,
                                  const char *target);
void lxw_add_ms_package_relationship(lxw_relationships *self,
                                     const char *type, const char *target);
void lxw_add_worksheet_relationship(lxw_relationships *self, const char *type,
                                    const char *target,
                                    const char *target_mode);
# 59 "/usr/local/include/xlsxwriter/worksheet.h" 2
# 79 "/usr/local/include/xlsxwriter/worksheet.h"
enum lxw_gridlines {

    LXW_HIDE_ALL_GRIDLINES = 0,


    LXW_SHOW_SCREEN_GRIDLINES,


    LXW_SHOW_PRINT_GRIDLINES,


    LXW_SHOW_ALL_GRIDLINES
};


enum lxw_validation_boolean {
    LXW_VALIDATION_DEFAULT,


    LXW_VALIDATION_OFF,



    LXW_VALIDATION_ON
};


enum lxw_validation_types {
    LXW_VALIDATION_TYPE_NONE,


    LXW_VALIDATION_TYPE_INTEGER,



    LXW_VALIDATION_TYPE_INTEGER_FORMULA,


    LXW_VALIDATION_TYPE_DECIMAL,



    LXW_VALIDATION_TYPE_DECIMAL_FORMULA,


    LXW_VALIDATION_TYPE_LIST,



    LXW_VALIDATION_TYPE_LIST_FORMULA,


    LXW_VALIDATION_TYPE_DATE,


    LXW_VALIDATION_TYPE_DATE_FORMULA,



    LXW_VALIDATION_TYPE_DATE_NUMBER,


    LXW_VALIDATION_TYPE_TIME,


    LXW_VALIDATION_TYPE_TIME_FORMULA,



    LXW_VALIDATION_TYPE_TIME_NUMBER,



    LXW_VALIDATION_TYPE_LENGTH,



    LXW_VALIDATION_TYPE_LENGTH_FORMULA,



    LXW_VALIDATION_TYPE_CUSTOM_FORMULA,


    LXW_VALIDATION_TYPE_ANY
};


enum lxw_validation_criteria {
    LXW_VALIDATION_CRITERIA_NONE,


    LXW_VALIDATION_CRITERIA_BETWEEN,


    LXW_VALIDATION_CRITERIA_NOT_BETWEEN,


    LXW_VALIDATION_CRITERIA_EQUAL_TO,


    LXW_VALIDATION_CRITERIA_NOT_EQUAL_TO,


    LXW_VALIDATION_CRITERIA_GREATER_THAN,


    LXW_VALIDATION_CRITERIA_LESS_THAN,


    LXW_VALIDATION_CRITERIA_GREATER_THAN_OR_EQUAL_TO,


    LXW_VALIDATION_CRITERIA_LESS_THAN_OR_EQUAL_TO
};


enum lxw_validation_error_types {

    LXW_VALIDATION_ERROR_TYPE_STOP,


    LXW_VALIDATION_ERROR_TYPE_WARNING,


    LXW_VALIDATION_ERROR_TYPE_INFORMATION
};



enum lxw_comment_display_types {

    LXW_COMMENT_DISPLAY_DEFAULT,


    LXW_COMMENT_DISPLAY_HIDDEN,



    LXW_COMMENT_DISPLAY_VISIBLE
};



enum lxw_object_position {


    LXW_OBJECT_POSITION_DEFAULT,


    LXW_OBJECT_MOVE_AND_SIZE,


    LXW_OBJECT_MOVE_DONT_SIZE,


    LXW_OBJECT_DONT_MOVE_DONT_SIZE,



    LXW_OBJECT_MOVE_AND_SIZE_AFTER
};

enum cell_types {
    NUMBER_CELL = 1,
    STRING_CELL,
    INLINE_STRING_CELL,
    INLINE_RICH_STRING_CELL,
    FORMULA_CELL,
    ARRAY_FORMULA_CELL,
    BLANK_CELL,
    BOOLEAN_CELL,
    COMMENT,
    HYPERLINK_URL,
    HYPERLINK_INTERNAL,
    HYPERLINK_EXTERNAL
};

enum pane_types {
    NO_PANES = 0,
    FREEZE_PANES,
    SPLIT_PANES,
    FREEZE_SPLIT_PANES
};


struct lxw_table_cells { struct lxw_cell *rbh_root; };
struct lxw_drawing_rel_ids { struct lxw_drawing_rel_id *rbh_root; };


struct lxw_table_rows {
    struct lxw_row *rbh_root;
    struct lxw_row *cached_row;
    lxw_row_t cached_row_num;
};
# 310 "/usr/local/include/xlsxwriter/worksheet.h"
struct lxw_merged_ranges { struct lxw_merged_range *stqh_first; struct lxw_merged_range **stqh_last; };
struct lxw_selections { struct lxw_selection *stqh_first; struct lxw_selection **stqh_last; };
struct lxw_data_validations { struct lxw_data_val_obj *stqh_first; struct lxw_data_val_obj **stqh_last; };
struct lxw_image_props { struct lxw_object_properties *stqh_first; struct lxw_object_properties **stqh_last; };
struct lxw_chart_props { struct lxw_object_properties *stqh_first; struct lxw_object_properties **stqh_last; };
struct lxw_comment_objs { struct lxw_vml_obj *stqh_first; struct lxw_vml_obj **stqh_last; };
# 332 "/usr/local/include/xlsxwriter/worksheet.h"
typedef struct lxw_row_col_options {

    uint8_t hidden;


    uint8_t level;


    uint8_t collapsed;
} lxw_row_col_options;

typedef struct lxw_col_options {
    lxw_col_t firstcol;
    lxw_col_t lastcol;
    double width;
    lxw_format *format;
    uint8_t hidden;
    uint8_t level;
    uint8_t collapsed;
} lxw_col_options;

typedef struct lxw_merged_range {
    lxw_row_t first_row;
    lxw_row_t last_row;
    lxw_col_t first_col;
    lxw_col_t last_col;

    struct { struct lxw_merged_range *stqe_next; } list_pointers;
} lxw_merged_range;

typedef struct lxw_repeat_rows {
    uint8_t in_use;
    lxw_row_t first_row;
    lxw_row_t last_row;
} lxw_repeat_rows;

typedef struct lxw_repeat_cols {
    uint8_t in_use;
    lxw_col_t first_col;
    lxw_col_t last_col;
} lxw_repeat_cols;

typedef struct lxw_print_area {
    uint8_t in_use;
    lxw_row_t first_row;
    lxw_row_t last_row;
    lxw_col_t first_col;
    lxw_col_t last_col;
} lxw_print_area;

typedef struct lxw_autofilter {
    uint8_t in_use;
    lxw_row_t first_row;
    lxw_row_t last_row;
    lxw_col_t first_col;
    lxw_col_t last_col;
} lxw_autofilter;

typedef struct lxw_panes {
    uint8_t type;
    lxw_row_t first_row;
    lxw_col_t first_col;
    lxw_row_t top_row;
    lxw_col_t left_col;
    double x_split;
    double y_split;
} lxw_panes;

typedef struct lxw_selection {
    char pane[12];
    char active_cell[(sizeof("$XFWD$1048576") * 2)];
    char sqref[(sizeof("$XFWD$1048576") * 2)];

    struct { struct lxw_selection *stqe_next; } list_pointers;

} lxw_selection;




typedef struct lxw_data_validation {




    uint8_t validate;





    uint8_t criteria;





    uint8_t ignore_blank;
# 438 "/usr/local/include/xlsxwriter/worksheet.h"
    uint8_t show_input;
# 447 "/usr/local/include/xlsxwriter/worksheet.h"
    uint8_t show_error;





    uint8_t error_type;







    uint8_t dropdown;





    double value_number;






    char *value_formula;
# 493 "/usr/local/include/xlsxwriter/worksheet.h"
    char **value_list;





    lxw_datetime value_datetime;





    double minimum_number;





    char *minimum_formula;





    lxw_datetime minimum_datetime;





    double maximum_number;





    char *maximum_formula;





    lxw_datetime maximum_datetime;
# 545 "/usr/local/include/xlsxwriter/worksheet.h"
    char *input_title;
# 554 "/usr/local/include/xlsxwriter/worksheet.h"
    char *input_message;







    char *error_title;
# 573 "/usr/local/include/xlsxwriter/worksheet.h"
    char *error_message;

} lxw_data_validation;




typedef struct lxw_data_val_obj {
    uint8_t validate;
    uint8_t criteria;
    uint8_t ignore_blank;
    uint8_t show_input;
    uint8_t show_error;
    uint8_t error_type;
    uint8_t dropdown;
    double value_number;
    char *value_formula;
    char **value_list;
    double minimum_number;
    char *minimum_formula;
    lxw_datetime minimum_datetime;
    double maximum_number;
    char *maximum_formula;
    lxw_datetime maximum_datetime;
    char *input_title;
    char *input_message;
    char *error_title;
    char *error_message;
    char sqref[(sizeof("$XFWD$1048576") * 2)];

    struct { struct lxw_data_val_obj *stqe_next; } list_pointers;
} lxw_data_val_obj;







typedef struct lxw_image_options {


    int32_t x_offset;


    int32_t y_offset;


    double x_scale;


    double y_scale;



    uint8_t object_position;



    char *description;



    char *url;


    char *tip;

} lxw_image_options;







typedef struct lxw_chart_options {


    int32_t x_offset;


    int32_t y_offset;


    double x_scale;


    double y_scale;



    uint8_t object_position;

} lxw_chart_options;




typedef struct lxw_object_properties {
    int32_t x_offset;
    int32_t y_offset;
    double x_scale;
    double y_scale;
    lxw_row_t row;
    lxw_col_t col;
    char *filename;
    char *description;
    char *url;
    char *tip;
    uint8_t object_position;
    FILE *stream;
    uint8_t image_type;
    uint8_t is_image_buffer;
    unsigned char *image_buffer;
    size_t image_buffer_size;
    double width;
    double height;
    char *extension;
    double x_dpi;
    double y_dpi;
    lxw_chart *chart;
    uint8_t is_duplicate;
    char *md5;

    struct { struct lxw_object_properties *stqe_next; } list_pointers;
} lxw_object_properties;







typedef struct lxw_comment_options {
# 716 "/usr/local/include/xlsxwriter/worksheet.h"
    uint8_t visible;






    char *author;




    uint16_t width;




    uint16_t height;



    double x_scale;



    double y_scale;




    lxw_color_t color;



    char *font_name;



    double font_size;



    uint8_t font_family;






    lxw_row_t start_row;




    lxw_col_t start_col;



    int32_t x_offset;



    int32_t y_offset;

} lxw_comment_options;


typedef struct lxw_vml_obj {

    lxw_row_t row;
    lxw_col_t col;
    lxw_row_t start_row;
    lxw_col_t start_col;
    int32_t x_offset;
    int32_t y_offset;
    uint32_t col_absolute;
    uint32_t row_absolute;
    uint32_t width;
    uint32_t height;
    lxw_color_t color;
    uint8_t font_family;
    uint8_t visible;
    uint32_t author_id;
    double font_size;
    struct lxw_drawing_coords from;
    struct lxw_drawing_coords to;
    char *author;
    char *font_name;
    char *text;
    struct { struct lxw_vml_obj *stqe_next; } list_pointers;

} lxw_vml_obj;
# 816 "/usr/local/include/xlsxwriter/worksheet.h"
typedef struct lxw_header_footer_options {

    double margin;
} lxw_header_footer_options;




typedef struct lxw_protection {

    uint8_t no_select_locked_cells;


    uint8_t no_select_unlocked_cells;


    uint8_t format_cells;


    uint8_t format_columns;


    uint8_t format_rows;


    uint8_t insert_columns;


    uint8_t insert_rows;


    uint8_t insert_hyperlinks;


    uint8_t delete_columns;


    uint8_t delete_rows;


    uint8_t sort;


    uint8_t autofilter;


    uint8_t pivot_tables;


    uint8_t scenarios;


    uint8_t objects;


    uint8_t no_content;


    uint8_t no_objects;

} lxw_protection;


typedef struct lxw_protection_obj {
    uint8_t no_select_locked_cells;
    uint8_t no_select_unlocked_cells;
    uint8_t format_cells;
    uint8_t format_columns;
    uint8_t format_rows;
    uint8_t insert_columns;
    uint8_t insert_rows;
    uint8_t insert_hyperlinks;
    uint8_t delete_columns;
    uint8_t delete_rows;
    uint8_t sort;
    uint8_t autofilter;
    uint8_t pivot_tables;
    uint8_t scenarios;
    uint8_t objects;
    uint8_t no_content;
    uint8_t no_objects;
    uint8_t no_sheet;
    uint8_t is_configured;
    char hash[5];
} lxw_protection_obj;
# 911 "/usr/local/include/xlsxwriter/worksheet.h"
typedef struct lxw_rich_string_tuple {



    lxw_format *format;


    char *string;
} lxw_rich_string_tuple;
# 928 "/usr/local/include/xlsxwriter/worksheet.h"
typedef struct lxw_worksheet {

    FILE *file;
    FILE *optimize_tmpfile;
    struct lxw_table_rows *table;
    struct lxw_table_rows *hyperlinks;
    struct lxw_table_rows *comments;
    struct lxw_cell **array;
    struct lxw_merged_ranges *merged_ranges;
    struct lxw_selections *selections;
    struct lxw_data_validations *data_validations;
    struct lxw_image_props *image_props;
    struct lxw_chart_props *chart_data;
    struct lxw_drawing_rel_ids *drawing_rel_ids;
    struct lxw_comment_objs *comment_objs;

    lxw_row_t dim_rowmin;
    lxw_row_t dim_rowmax;
    lxw_col_t dim_colmin;
    lxw_col_t dim_colmax;

    lxw_sst *sst;
    char *name;
    char *quoted_name;
    char *tmpdir;

    uint32_t index;
    uint8_t active;
    uint8_t selected;
    uint8_t hidden;
    uint16_t *active_sheet;
    uint16_t *first_sheet;
    uint8_t is_chartsheet;

    lxw_col_options **col_options;
    uint16_t col_options_max;

    double *col_sizes;
    uint16_t col_sizes_max;

    lxw_format **col_formats;
    uint16_t col_formats_max;

    uint8_t col_size_changed;
    uint8_t row_size_changed;
    uint8_t optimize;
    struct lxw_row *optimize_row;

    uint16_t fit_height;
    uint16_t fit_width;
    uint16_t horizontal_dpi;
    uint16_t hlink_count;
    uint16_t page_start;
    uint16_t print_scale;
    uint16_t rel_count;
    uint16_t vertical_dpi;
    uint16_t zoom;
    uint8_t filter_on;
    uint8_t fit_page;
    uint8_t hcenter;
    uint8_t orientation;
    uint8_t outline_changed;
    uint8_t outline_on;
    uint8_t outline_style;
    uint8_t outline_below;
    uint8_t outline_right;
    uint8_t page_order;
    uint8_t page_setup_changed;
    uint8_t page_view;
    uint8_t paper_size;
    uint8_t print_gridlines;
    uint8_t print_headers;
    uint8_t print_options_changed;
    uint8_t right_to_left;
    uint8_t screen_gridlines;
    uint8_t show_zeros;
    uint8_t vcenter;
    uint8_t zoom_scale_normal;
    uint8_t num_validations;
    char *vba_codename;

    lxw_color_t tab_color;

    double margin_left;
    double margin_right;
    double margin_top;
    double margin_bottom;
    double margin_header;
    double margin_footer;

    double default_row_height;
    uint32_t default_row_pixels;
    uint32_t default_col_pixels;
    uint8_t default_row_zeroed;
    uint8_t default_row_set;
    uint8_t outline_row_level;
    uint8_t outline_col_level;

    uint8_t header_footer_changed;
    char header[255];
    char footer[255];

    struct lxw_repeat_rows repeat_rows;
    struct lxw_repeat_cols repeat_cols;
    struct lxw_print_area print_area;
    struct lxw_autofilter autofilter;

    uint16_t merged_range_count;
    uint16_t max_url_length;

    lxw_row_t *hbreaks;
    lxw_col_t *vbreaks;
    uint16_t hbreaks_count;
    uint16_t vbreaks_count;

    uint32_t drawing_rel_id;
    struct lxw_rel_tuples *external_hyperlinks;
    struct lxw_rel_tuples *external_drawing_links;
    struct lxw_rel_tuples *drawing_links;

    struct lxw_panes panes;

    struct lxw_protection_obj protection;

    lxw_drawing *drawing;
    lxw_format *default_url_format;

    uint8_t has_vml;
    uint8_t has_comments;
    uint8_t has_header_vml;
    lxw_rel_tuple *external_vml_comment_link;
    lxw_rel_tuple *external_comment_link;
    char *comment_author;
    char *vml_data_id_str;
    uint32_t vml_shape_id;
    uint8_t comment_display_default;

    struct { struct lxw_worksheet *stqe_next; } list_pointers;

} lxw_worksheet;




typedef struct lxw_worksheet_init_data {
    uint32_t index;
    uint8_t hidden;
    uint8_t optimize;
    uint16_t *active_sheet;
    uint16_t *first_sheet;
    lxw_sst *sst;
    char *name;
    char *quoted_name;
    char *tmpdir;
    lxw_format *default_url_format;
    uint16_t max_url_length;

} lxw_worksheet_init_data;


typedef struct lxw_row {
    lxw_row_t row_num;
    double height;
    lxw_format *format;
    uint8_t hidden;
    uint8_t level;
    uint8_t collapsed;
    uint8_t row_changed;
    uint8_t data_changed;
    uint8_t height_changed;

    struct lxw_table_cells *cells;


    struct { struct lxw_row *rbe_left; struct lxw_row *rbe_right; struct lxw_row *rbe_parent; int rbe_color; } tree_pointers;
} lxw_row;


typedef struct lxw_cell {
    lxw_row_t row_num;
    lxw_col_t col_num;
    enum cell_types type;
    lxw_format *format;
    lxw_vml_obj *comment;

    union {
        double number;
        int32_t string_id;
        char *string;
    } u;

    double formula_result;
    char *user_data1;
    char *user_data2;
    char *sst_string;


    struct { struct lxw_cell *rbe_left; struct lxw_cell *rbe_right; struct lxw_cell *rbe_parent; int rbe_color; } tree_pointers;
} lxw_cell;


typedef struct lxw_drawing_rel_id {
    uint32_t id;
    char *target;

    struct { struct lxw_drawing_rel_id *rbe_left; struct lxw_drawing_rel_id *rbe_right; struct lxw_drawing_rel_id *rbe_parent; int rbe_color; } tree_pointers;
} lxw_drawing_rel_id;
# 1185 "/usr/local/include/xlsxwriter/worksheet.h"
lxw_error worksheet_write_number(lxw_worksheet *worksheet,
                                 lxw_row_t row,
                                 lxw_col_t col, double number,
                                 lxw_format *format);
# 1233 "/usr/local/include/xlsxwriter/worksheet.h"
lxw_error worksheet_write_string(lxw_worksheet *worksheet,
                                 lxw_row_t row,
                                 lxw_col_t col, const char *string,
                                 lxw_format *format);
# 1286 "/usr/local/include/xlsxwriter/worksheet.h"
lxw_error worksheet_write_formula(lxw_worksheet *worksheet,
                                  lxw_row_t row,
                                  lxw_col_t col, const char *formula,
                                  lxw_format *format);
# 1332 "/usr/local/include/xlsxwriter/worksheet.h"
lxw_error worksheet_write_array_formula(lxw_worksheet *worksheet,
                                        lxw_row_t first_row,
                                        lxw_col_t first_col,
                                        lxw_row_t last_row,
                                        lxw_col_t last_col,
                                        const char *formula,
                                        lxw_format *format);

lxw_error worksheet_write_array_formula_num(lxw_worksheet *worksheet,
                                            lxw_row_t first_row,
                                            lxw_col_t first_col,
                                            lxw_row_t last_row,
                                            lxw_col_t last_col,
                                            const char *formula,
                                            lxw_format *format,
                                            double result);
# 1376 "/usr/local/include/xlsxwriter/worksheet.h"
lxw_error worksheet_write_datetime(lxw_worksheet *worksheet,
                                   lxw_row_t row,
                                   lxw_col_t col, lxw_datetime *datetime,
                                   lxw_format *format);
# 1521 "/usr/local/include/xlsxwriter/worksheet.h"
lxw_error worksheet_write_url(lxw_worksheet *worksheet,
                              lxw_row_t row,
                              lxw_col_t col, const char *url,
                              lxw_format *format);




lxw_error worksheet_write_url_opt(lxw_worksheet *worksheet,
                                  lxw_row_t row_num,
                                  lxw_col_t col_num, const char *url,
                                  lxw_format *format, const char *string,
                                  const char *tooltip);
# 1553 "/usr/local/include/xlsxwriter/worksheet.h"
lxw_error worksheet_write_boolean(lxw_worksheet *worksheet,
                                  lxw_row_t row, lxw_col_t col,
                                  int value, lxw_format *format);
# 1584 "/usr/local/include/xlsxwriter/worksheet.h"
lxw_error worksheet_write_blank(lxw_worksheet *worksheet,
                                lxw_row_t row, lxw_col_t col,
                                lxw_format *format);
# 1631 "/usr/local/include/xlsxwriter/worksheet.h"
lxw_error worksheet_write_formula_num(lxw_worksheet *worksheet,
                                      lxw_row_t row,
                                      lxw_col_t col,
                                      const char *formula,
                                      lxw_format *format, double result);
# 1706 "/usr/local/include/xlsxwriter/worksheet.h"
lxw_error worksheet_write_rich_string(lxw_worksheet *worksheet,
                                      lxw_row_t row,
                                      lxw_col_t col,
                                      lxw_rich_string_tuple *rich_string[],
                                      lxw_format *format);
# 1738 "/usr/local/include/xlsxwriter/worksheet.h"
lxw_error worksheet_write_comment(lxw_worksheet *worksheet,
                                  lxw_row_t row, lxw_col_t col,
                                  const char *string);
# 1789 "/usr/local/include/xlsxwriter/worksheet.h"
lxw_error worksheet_write_comment_opt(lxw_worksheet *worksheet,
                                      lxw_row_t row, lxw_col_t col,
                                      const char *string,
                                      lxw_comment_options *options);
# 1846 "/usr/local/include/xlsxwriter/worksheet.h"
lxw_error worksheet_set_row(lxw_worksheet *worksheet,
                            lxw_row_t row, double height, lxw_format *format);
# 1908 "/usr/local/include/xlsxwriter/worksheet.h"
lxw_error worksheet_set_row_opt(lxw_worksheet *worksheet,
                                lxw_row_t row,
                                double height,
                                lxw_format *format,
                                lxw_row_col_options *options);
# 2009 "/usr/local/include/xlsxwriter/worksheet.h"
lxw_error worksheet_set_column(lxw_worksheet *worksheet,
                               lxw_col_t first_col,
                               lxw_col_t last_col,
                               double width, lxw_format *format);
# 2056 "/usr/local/include/xlsxwriter/worksheet.h"
lxw_error worksheet_set_column_opt(lxw_worksheet *worksheet,
                                   lxw_col_t first_col,
                                   lxw_col_t last_col,
                                   double width,
                                   lxw_format *format,
                                   lxw_row_col_options *options);
# 2096 "/usr/local/include/xlsxwriter/worksheet.h"
lxw_error worksheet_insert_image(lxw_worksheet *worksheet,
                                 lxw_row_t row, lxw_col_t col,
                                 const char *filename);
# 2150 "/usr/local/include/xlsxwriter/worksheet.h"
lxw_error worksheet_insert_image_opt(lxw_worksheet *worksheet,
                                     lxw_row_t row, lxw_col_t col,
                                     const char *filename,
                                     lxw_image_options *options);
# 2181 "/usr/local/include/xlsxwriter/worksheet.h"
lxw_error worksheet_insert_image_buffer(lxw_worksheet *worksheet,
                                        lxw_row_t row,
                                        lxw_col_t col,
                                        const unsigned char *image_buffer,
                                        size_t image_size);
# 2218 "/usr/local/include/xlsxwriter/worksheet.h"
lxw_error worksheet_insert_image_buffer_opt(lxw_worksheet *worksheet,
                                            lxw_row_t row,
                                            lxw_col_t col,
                                            const unsigned char *image_buffer,
                                            size_t image_size,
                                            lxw_image_options *options);
# 2260 "/usr/local/include/xlsxwriter/worksheet.h"
lxw_error worksheet_insert_chart(lxw_worksheet *worksheet,
                                 lxw_row_t row, lxw_col_t col,
                                 lxw_chart *chart);
# 2290 "/usr/local/include/xlsxwriter/worksheet.h"
lxw_error worksheet_insert_chart_opt(lxw_worksheet *worksheet,
                                     lxw_row_t row, lxw_col_t col,
                                     lxw_chart *chart,
                                     lxw_chart_options *user_options);
# 2355 "/usr/local/include/xlsxwriter/worksheet.h"
lxw_error worksheet_merge_range(lxw_worksheet *worksheet, lxw_row_t first_row,
                                lxw_col_t first_col, lxw_row_t last_row,
                                lxw_col_t last_col, const char *string,
                                lxw_format *format);
# 2392 "/usr/local/include/xlsxwriter/worksheet.h"
lxw_error worksheet_autofilter(lxw_worksheet *worksheet, lxw_row_t first_row,
                               lxw_col_t first_col, lxw_row_t last_row,
                               lxw_col_t last_col);
# 2431 "/usr/local/include/xlsxwriter/worksheet.h"
lxw_error worksheet_data_validation_cell(lxw_worksheet *worksheet,
                                         lxw_row_t row, lxw_col_t col,
                                         lxw_data_validation *validation);
# 2470 "/usr/local/include/xlsxwriter/worksheet.h"
lxw_error worksheet_data_validation_range(lxw_worksheet *worksheet,
                                          lxw_row_t first_row,
                                          lxw_col_t first_col,
                                          lxw_row_t last_row,
                                          lxw_col_t last_col,
                                          lxw_data_validation *validation);
# 2501 "/usr/local/include/xlsxwriter/worksheet.h"
void worksheet_activate(lxw_worksheet *worksheet);
# 2524 "/usr/local/include/xlsxwriter/worksheet.h"
void worksheet_select(lxw_worksheet *worksheet);
# 2553 "/usr/local/include/xlsxwriter/worksheet.h"
void worksheet_hide(lxw_worksheet *worksheet);
# 2573 "/usr/local/include/xlsxwriter/worksheet.h"
void worksheet_set_first_sheet(lxw_worksheet *worksheet);
# 2605 "/usr/local/include/xlsxwriter/worksheet.h"
void worksheet_freeze_panes(lxw_worksheet *worksheet,
                            lxw_row_t row, lxw_col_t col);
# 2638 "/usr/local/include/xlsxwriter/worksheet.h"
void worksheet_split_panes(lxw_worksheet *worksheet,
                           double vertical, double horizontal);


void worksheet_freeze_panes_opt(lxw_worksheet *worksheet,
                                lxw_row_t first_row, lxw_col_t first_col,
                                lxw_row_t top_row, lxw_col_t left_col,
                                uint8_t type);


void worksheet_split_panes_opt(lxw_worksheet *worksheet,
                               double vertical, double horizontal,
                               lxw_row_t top_row, lxw_col_t left_col);
# 2680 "/usr/local/include/xlsxwriter/worksheet.h"
void worksheet_set_selection(lxw_worksheet *worksheet,
                             lxw_row_t first_row, lxw_col_t first_col,
                             lxw_row_t last_row, lxw_col_t last_col);
# 2696 "/usr/local/include/xlsxwriter/worksheet.h"
void worksheet_set_landscape(lxw_worksheet *worksheet);
# 2711 "/usr/local/include/xlsxwriter/worksheet.h"
void worksheet_set_portrait(lxw_worksheet *worksheet);
# 2724 "/usr/local/include/xlsxwriter/worksheet.h"
void worksheet_set_page_view(lxw_worksheet *worksheet);
# 2793 "/usr/local/include/xlsxwriter/worksheet.h"
void worksheet_set_paper(lxw_worksheet *worksheet, uint8_t paper_type);
# 2813 "/usr/local/include/xlsxwriter/worksheet.h"
void worksheet_set_margins(lxw_worksheet *worksheet, double left,
                           double right, double top, double bottom);
# 2995 "/usr/local/include/xlsxwriter/worksheet.h"
lxw_error worksheet_set_header(lxw_worksheet *worksheet, const char *string);
# 3008 "/usr/local/include/xlsxwriter/worksheet.h"
lxw_error worksheet_set_footer(lxw_worksheet *worksheet, const char *string);
# 3033 "/usr/local/include/xlsxwriter/worksheet.h"
lxw_error worksheet_set_header_opt(lxw_worksheet *worksheet,
                                   const char *string,
                                   lxw_header_footer_options *options);
# 3049 "/usr/local/include/xlsxwriter/worksheet.h"
lxw_error worksheet_set_footer_opt(lxw_worksheet *worksheet,
                                   const char *string,
                                   lxw_header_footer_options *options);
# 3093 "/usr/local/include/xlsxwriter/worksheet.h"
lxw_error worksheet_set_h_pagebreaks(lxw_worksheet *worksheet,
                                     lxw_row_t breaks[]);
# 3136 "/usr/local/include/xlsxwriter/worksheet.h"
lxw_error worksheet_set_v_pagebreaks(lxw_worksheet *worksheet,
                                     lxw_col_t breaks[]);
# 3164 "/usr/local/include/xlsxwriter/worksheet.h"
void worksheet_print_across(lxw_worksheet *worksheet);
# 3187 "/usr/local/include/xlsxwriter/worksheet.h"
void worksheet_set_zoom(lxw_worksheet *worksheet, uint16_t scale);
# 3209 "/usr/local/include/xlsxwriter/worksheet.h"
void worksheet_gridlines(lxw_worksheet *worksheet, uint8_t option);
# 3224 "/usr/local/include/xlsxwriter/worksheet.h"
void worksheet_center_horizontally(lxw_worksheet *worksheet);
# 3239 "/usr/local/include/xlsxwriter/worksheet.h"
void worksheet_center_vertically(lxw_worksheet *worksheet);
# 3258 "/usr/local/include/xlsxwriter/worksheet.h"
void worksheet_print_row_col_headers(lxw_worksheet *worksheet);
# 3280 "/usr/local/include/xlsxwriter/worksheet.h"
lxw_error worksheet_repeat_rows(lxw_worksheet *worksheet, lxw_row_t first_row,
                                lxw_row_t last_row);
# 3303 "/usr/local/include/xlsxwriter/worksheet.h"
lxw_error worksheet_repeat_columns(lxw_worksheet *worksheet,
                                   lxw_col_t first_col, lxw_col_t last_col);
# 3333 "/usr/local/include/xlsxwriter/worksheet.h"
lxw_error worksheet_print_area(lxw_worksheet *worksheet, lxw_row_t first_row,
                               lxw_col_t first_col, lxw_row_t last_row,
                               lxw_col_t last_col);
# 3383 "/usr/local/include/xlsxwriter/worksheet.h"
void worksheet_fit_to_pages(lxw_worksheet *worksheet, uint16_t width,
                            uint16_t height);
# 3400 "/usr/local/include/xlsxwriter/worksheet.h"
void worksheet_set_start_page(lxw_worksheet *worksheet, uint16_t start_page);
# 3426 "/usr/local/include/xlsxwriter/worksheet.h"
void worksheet_set_print_scale(lxw_worksheet *worksheet, uint16_t scale);
# 3445 "/usr/local/include/xlsxwriter/worksheet.h"
void worksheet_right_to_left(lxw_worksheet *worksheet);
# 3459 "/usr/local/include/xlsxwriter/worksheet.h"
void worksheet_hide_zero(lxw_worksheet *worksheet);
# 3478 "/usr/local/include/xlsxwriter/worksheet.h"
void worksheet_set_tab_color(lxw_worksheet *worksheet, lxw_color_t color);
# 3556 "/usr/local/include/xlsxwriter/worksheet.h"
void worksheet_protect(lxw_worksheet *worksheet, const char *password,
                       lxw_protection *options);
# 3603 "/usr/local/include/xlsxwriter/worksheet.h"
void worksheet_outline_settings(lxw_worksheet *worksheet, uint8_t visible,
                                uint8_t symbols_below, uint8_t symbols_right,
                                uint8_t auto_style);
# 3635 "/usr/local/include/xlsxwriter/worksheet.h"
void worksheet_set_default_row(lxw_worksheet *worksheet, double height,
                               uint8_t hide_unused_rows);
# 3662 "/usr/local/include/xlsxwriter/worksheet.h"
lxw_error worksheet_set_vba_name(lxw_worksheet *worksheet, const char *name);
# 3680 "/usr/local/include/xlsxwriter/worksheet.h"
void worksheet_show_comments(lxw_worksheet *worksheet);
# 3699 "/usr/local/include/xlsxwriter/worksheet.h"
void worksheet_set_comments_author(lxw_worksheet *worksheet,
                                   const char *author);

lxw_worksheet *lxw_worksheet_new(lxw_worksheet_init_data *init_data);
void lxw_worksheet_free(lxw_worksheet *worksheet);
void lxw_worksheet_assemble_xml_file(lxw_worksheet *worksheet);
void lxw_worksheet_write_single_row(lxw_worksheet *worksheet);

void lxw_worksheet_prepare_image(lxw_worksheet *worksheet,
                                 uint32_t image_ref_id, uint32_t drawing_id,
                                 lxw_object_properties *object_props);

void lxw_worksheet_prepare_chart(lxw_worksheet *worksheet,
                                 uint32_t chart_ref_id, uint32_t drawing_id,
                                 lxw_object_properties *object_props,
                                 uint8_t is_chartsheet);

uint32_t lxw_worksheet_prepare_vml_objects(lxw_worksheet *worksheet,
                                           uint32_t vml_data_id,
                                           uint32_t vml_shape_id,
                                           uint32_t vml_drawing_id,
                                           uint32_t comment_id);

lxw_row *lxw_worksheet_find_row(lxw_worksheet *worksheet, lxw_row_t row_num);
lxw_cell *lxw_worksheet_find_cell_in_row(lxw_row *row, lxw_col_t col_num);



void lxw_worksheet_write_sheet_views(lxw_worksheet *worksheet);
void lxw_worksheet_write_page_margins(lxw_worksheet *worksheet);
void lxw_worksheet_write_drawings(lxw_worksheet *worksheet);
void lxw_worksheet_write_sheet_protection(lxw_worksheet *worksheet,
                                          lxw_protection_obj *protect);
void lxw_worksheet_write_sheet_pr(lxw_worksheet *worksheet);
void lxw_worksheet_write_page_setup(lxw_worksheet *worksheet);
void lxw_worksheet_write_header_footer(lxw_worksheet *worksheet);
# 49 "/usr/local/include/xlsxwriter/workbook.h" 2
# 1 "/usr/local/include/xlsxwriter/chartsheet.h" 1
# 74 "/usr/local/include/xlsxwriter/chartsheet.h"
typedef struct lxw_chartsheet {

    FILE *file;
    lxw_worksheet *worksheet;
    lxw_chart *chart;

    struct lxw_protection_obj protection;
    uint8_t is_protected;

    char *name;
    char *quoted_name;
    char *tmpdir;
    uint32_t index;
    uint8_t active;
    uint8_t selected;
    uint8_t hidden;
    uint16_t *active_sheet;
    uint16_t *first_sheet;
    uint16_t rel_count;

    struct { struct lxw_chartsheet *stqe_next; } list_pointers;

} lxw_chartsheet;
# 141 "/usr/local/include/xlsxwriter/chartsheet.h"
lxw_error chartsheet_set_chart(lxw_chartsheet *chartsheet, lxw_chart *chart);


lxw_error chartsheet_set_chart_opt(lxw_chartsheet *chartsheet,
                                   lxw_chart *chart,
                                   lxw_chart_options *user_options);
# 175 "/usr/local/include/xlsxwriter/chartsheet.h"
void chartsheet_activate(lxw_chartsheet *chartsheet);
# 200 "/usr/local/include/xlsxwriter/chartsheet.h"
void chartsheet_select(lxw_chartsheet *chartsheet);
# 232 "/usr/local/include/xlsxwriter/chartsheet.h"
void chartsheet_hide(lxw_chartsheet *chartsheet);
# 256 "/usr/local/include/xlsxwriter/chartsheet.h"
void chartsheet_set_first_sheet(lxw_chartsheet *chartsheet);
# 277 "/usr/local/include/xlsxwriter/chartsheet.h"
void chartsheet_set_tab_color(lxw_chartsheet *chartsheet, lxw_color_t color);
# 332 "/usr/local/include/xlsxwriter/chartsheet.h"
void chartsheet_protect(lxw_chartsheet *chartsheet, const char *password,
                        lxw_protection *options);
# 352 "/usr/local/include/xlsxwriter/chartsheet.h"
void chartsheet_set_zoom(lxw_chartsheet *chartsheet, uint16_t scale);
# 367 "/usr/local/include/xlsxwriter/chartsheet.h"
void chartsheet_set_landscape(lxw_chartsheet *chartsheet);
# 381 "/usr/local/include/xlsxwriter/chartsheet.h"
void chartsheet_set_portrait(lxw_chartsheet *chartsheet);
# 402 "/usr/local/include/xlsxwriter/chartsheet.h"
void chartsheet_set_paper(lxw_chartsheet *chartsheet, uint8_t paper_type);
# 422 "/usr/local/include/xlsxwriter/chartsheet.h"
void chartsheet_set_margins(lxw_chartsheet *chartsheet, double left,
                            double right, double top, double bottom);
# 467 "/usr/local/include/xlsxwriter/chartsheet.h"
lxw_error chartsheet_set_header(lxw_chartsheet *chartsheet,
                                const char *string);
# 481 "/usr/local/include/xlsxwriter/chartsheet.h"
lxw_error chartsheet_set_footer(lxw_chartsheet *chartsheet,
                                const char *string);
# 507 "/usr/local/include/xlsxwriter/chartsheet.h"
lxw_error chartsheet_set_header_opt(lxw_chartsheet *chartsheet,
                                    const char *string,
                                    lxw_header_footer_options *options);
# 523 "/usr/local/include/xlsxwriter/chartsheet.h"
lxw_error chartsheet_set_footer_opt(lxw_chartsheet *chartsheet,
                                    const char *string,
                                    lxw_header_footer_options *options);

lxw_chartsheet *lxw_chartsheet_new(lxw_worksheet_init_data *init_data);
void lxw_chartsheet_free(lxw_chartsheet *chartsheet);
void lxw_chartsheet_assemble_xml_file(lxw_chartsheet *chartsheet);
# 50 "/usr/local/include/xlsxwriter/workbook.h" 2
# 58 "/usr/local/include/xlsxwriter/workbook.h"
struct lxw_worksheet_names { struct lxw_worksheet_name *rbh_root; };
struct lxw_chartsheet_names { struct lxw_chartsheet_name *rbh_root; };
struct lxw_image_md5s { struct lxw_image_md5 *rbh_root; };


struct lxw_sheets { struct lxw_sheet *stqh_first; struct lxw_sheet **stqh_last; };
struct lxw_worksheets { struct lxw_worksheet *stqh_first; struct lxw_worksheet **stqh_last; };
struct lxw_chartsheets { struct lxw_chartsheet *stqh_first; struct lxw_chartsheet **stqh_last; };
struct lxw_charts { struct lxw_chart *stqh_first; struct lxw_chart **stqh_last; };
struct lxw_defined_names { struct lxw_defined_name *tqh_first; struct lxw_defined_name **tqh_last; };


typedef struct lxw_sheet {
    uint8_t is_chartsheet;

    union {
        lxw_worksheet *worksheet;
        lxw_chartsheet *chartsheet;
    } u;

    struct { struct lxw_sheet *stqe_next; } list_pointers;
} lxw_sheet;


typedef struct lxw_worksheet_name {
    const char *name;
    lxw_worksheet *worksheet;

    struct { struct lxw_worksheet_name *rbe_left; struct lxw_worksheet_name *rbe_right; struct lxw_worksheet_name *rbe_parent; int rbe_color; } tree_pointers;
} lxw_worksheet_name;


typedef struct lxw_chartsheet_name {
    const char *name;
    lxw_chartsheet *chartsheet;

    struct { struct lxw_chartsheet_name *rbe_left; struct lxw_chartsheet_name *rbe_right; struct lxw_chartsheet_name *rbe_parent; int rbe_color; } tree_pointers;
} lxw_chartsheet_name;


typedef struct lxw_image_md5 {
    uint32_t id;
    char *md5;

    struct { struct lxw_image_md5 *rbe_left; struct lxw_image_md5 *rbe_right; struct lxw_image_md5 *rbe_parent; int rbe_color; } tree_pointers;
} lxw_image_md5;
# 167 "/usr/local/include/xlsxwriter/workbook.h"
typedef struct lxw_defined_name {
    int16_t index;
    uint8_t hidden;
    char name[128];
    char app_name[128];
    char formula[128];
    char normalised_name[128];
    char normalised_sheetname[128];


    struct { struct lxw_defined_name *tqe_next; struct lxw_defined_name **tqe_prev; } list_pointers;
} lxw_defined_name;




typedef struct lxw_doc_properties {

    char *title;


    char *subject;


    char *author;


    char *manager;


    char *company;


    char *category;


    char *keywords;


    char *comments;


    char *status;


    char *hyperlink_base;





    time_t created;

} lxw_doc_properties;
# 254 "/usr/local/include/xlsxwriter/workbook.h"
typedef struct lxw_workbook_options {

    uint8_t constant_memory;


    char *tmpdir;


    uint8_t use_zip64;

} lxw_workbook_options;
# 273 "/usr/local/include/xlsxwriter/workbook.h"
typedef struct lxw_workbook {

    FILE *file;
    struct lxw_sheets *sheets;
    struct lxw_worksheets *worksheets;
    struct lxw_chartsheets *chartsheets;
    struct lxw_worksheet_names *worksheet_names;
    struct lxw_chartsheet_names *chartsheet_names;
    struct lxw_image_md5s *image_md5s;
    struct lxw_charts *charts;
    struct lxw_charts *ordered_charts;
    struct lxw_formats *formats;
    struct lxw_defined_names *defined_names;
    lxw_sst *sst;
    lxw_doc_properties *properties;
    struct lxw_custom_properties *custom_properties;

    char *filename;
    lxw_workbook_options options;

    uint16_t num_sheets;
    uint16_t num_worksheets;
    uint16_t num_chartsheets;
    uint16_t first_sheet;
    uint16_t active_sheet;
    uint16_t num_xf_formats;
    uint16_t num_format_count;
    uint16_t drawing_count;
    uint16_t comment_count;

    uint16_t font_count;
    uint16_t border_count;
    uint16_t fill_count;
    uint8_t optimize;
    uint16_t max_url_length;

    uint8_t has_png;
    uint8_t has_jpeg;
    uint8_t has_bmp;
    uint8_t has_vml;
    uint8_t has_comments;

    lxw_hash_table *used_xf_formats;

    char *vba_project;
    char *vba_codename;

    lxw_format *default_url_format;

} lxw_workbook;
# 349 "/usr/local/include/xlsxwriter/workbook.h"
lxw_workbook *workbook_new(const char *filename);
# 396 "/usr/local/include/xlsxwriter/workbook.h"
lxw_workbook *workbook_new_opt(const char *filename,
                               lxw_workbook_options *options);
# 438 "/usr/local/include/xlsxwriter/workbook.h"
lxw_worksheet *workbook_add_worksheet(lxw_workbook *workbook,
                                      const char *sheetname);
# 482 "/usr/local/include/xlsxwriter/workbook.h"
lxw_chartsheet *workbook_add_chartsheet(lxw_workbook *workbook,
                                        const char *sheetname);
# 512 "/usr/local/include/xlsxwriter/workbook.h"
lxw_format *workbook_add_format(lxw_workbook *workbook);
# 568 "/usr/local/include/xlsxwriter/workbook.h"
lxw_chart *workbook_add_chart(lxw_workbook *workbook, uint8_t chart_type);
# 594 "/usr/local/include/xlsxwriter/workbook.h"
lxw_error workbook_close(lxw_workbook *workbook);
# 654 "/usr/local/include/xlsxwriter/workbook.h"
lxw_error workbook_set_properties(lxw_workbook *workbook,
                                  lxw_doc_properties *properties);
# 690 "/usr/local/include/xlsxwriter/workbook.h"
lxw_error workbook_set_custom_property_string(lxw_workbook *workbook,
                                              const char *name,
                                              const char *value);
# 709 "/usr/local/include/xlsxwriter/workbook.h"
lxw_error workbook_set_custom_property_number(lxw_workbook *workbook,
                                              const char *name, double value);




lxw_error workbook_set_custom_property_integer(lxw_workbook *workbook,
                                               const char *name,
                                               int32_t value);
# 735 "/usr/local/include/xlsxwriter/workbook.h"
lxw_error workbook_set_custom_property_boolean(lxw_workbook *workbook,
                                               const char *name,
                                               uint8_t value);
# 756 "/usr/local/include/xlsxwriter/workbook.h"
lxw_error workbook_set_custom_property_datetime(lxw_workbook *workbook,
                                                const char *name,
                                                lxw_datetime *datetime);
# 808 "/usr/local/include/xlsxwriter/workbook.h"
lxw_error workbook_define_name(lxw_workbook *workbook, const char *name,
                               const char *formula);
# 829 "/usr/local/include/xlsxwriter/workbook.h"
lxw_format *workbook_get_default_url_format(lxw_workbook *workbook);
# 846 "/usr/local/include/xlsxwriter/workbook.h"
lxw_worksheet *workbook_get_worksheet_by_name(lxw_workbook *workbook,
                                              const char *name);
# 864 "/usr/local/include/xlsxwriter/workbook.h"
lxw_chartsheet *workbook_get_chartsheet_by_name(lxw_workbook *workbook,
                                                const char *name);
# 899 "/usr/local/include/xlsxwriter/workbook.h"
lxw_error workbook_validate_sheet_name(lxw_workbook *workbook,
                                       const char *sheetname);
# 932 "/usr/local/include/xlsxwriter/workbook.h"
lxw_error workbook_add_vba_project(lxw_workbook *workbook,
                                   const char *filename);
# 957 "/usr/local/include/xlsxwriter/workbook.h"
lxw_error workbook_set_vba_name(lxw_workbook *workbook, const char *name);

void lxw_workbook_free(lxw_workbook *workbook);
void lxw_workbook_assemble_xml_file(lxw_workbook *workbook);
void lxw_workbook_set_default_xf_indices(lxw_workbook *workbook);
void workbook_unset_default_url_format(lxw_workbook *workbook);
# 17 "/usr/local/include/xlsxwriter.h" 2
# 1 "/usr/local/include/xlsxwriter/worksheet.h" 1
# 18 "/usr/local/include/xlsxwriter.h" 2
# 1 "/usr/local/include/xlsxwriter/format.h" 1
# 19 "/usr/local/include/xlsxwriter.h" 2
# 1 "/usr/local/include/xlsxwriter/utility.h" 1
# 20 "/usr/local/include/xlsxwriter.h" 2
